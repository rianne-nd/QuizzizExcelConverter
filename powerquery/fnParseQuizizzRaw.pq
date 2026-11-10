// ===========================================================================
//  fnParseQuizizzRaw
//  Parses ONE Quizizz "raw" .txt file into Quizizz-template rows.
//
//  Query type : Function
//  Signature  : (Source as binary) as table
//
//  It takes the file's CONTENT (a binary), not a path. That way the caller
//  holds the File.Contents / Folder.Files step, which lets you click the gear
//  icon in Applied Steps to browse for a different file instead of editing
//  the path by hand. It also avoids reading the same file twice in folder mode.
//
//  WHY THIS EXISTS
//  The old query built RecordID with
//      each if Text.StartsWith([Column1], "Question:") then 1 else null
//  which sets 1 for EVERY question - it never increments. Grouping on a
//  constant therefore merged any two questions that happened to share the
//  same question text, producing 8 option rows for a 4-option question,
//  a phantom option 5, and answer keys like "4,8". That silently destroyed
//  intentionally repeated questions (used as recall items). This version
//  anchors each block to a real, incrementing question number, so blocks are
//  never merged and every repeat stays its own question.
//
//  It also fixes:
//    * only the TRAILING "*" is stripped (inline emphasis like *before* survives)
//    * a missing "Explanation:" line no longer leaks the next question's one
//    * "Image:" lines are parsed into the Image Link column
//    * zero-width characters from Word/PDF copy-paste are removed
//    * options beyond the template's 5-column limit are dropped AND the answer
//      key is re-validated, instead of silently pointing at a dropped option
// ===========================================================================
(Source as binary) as table =>
let
    TimeSeconds = 60,
    MaxOptions  = 5,

    ZeroWidth = {
        Character.FromNumber(8203),   // zero-width space
        Character.FromNumber(8204),   // zero-width non-joiner
        Character.FromNumber(8205),   // zero-width joiner
        Character.FromNumber(65279)   // BOM / zero-width no-break space
    },

    // ---------- 1. read the file as clean, non-empty lines ----------
    RawLines   = Lines.FromBinary(Source, null, null, 65001),
    CleanLines = List.Transform(
        RawLines,
        each Text.Trim(Text.Replace(Text.Remove(_, ZeroWidth), "#(00A0)", " "))
    ),
    NonEmpty   = List.Select(CleanLines, each _ <> ""),
    AsTable    = Table.FromColumns({NonEmpty}, {"Line"}),

    // ---------- 2. classify every line ----------
    Classified = Table.AddColumn(AsTable, "Parts", each
        let t = [Line] in
            if      Text.StartsWith(t, "Question:",    Comparer.OrdinalIgnoreCase) then [Kind = "Question",    Body = Text.Trim(Text.Range(t, 9))]
            else if Text.StartsWith(t, "Option:",      Comparer.OrdinalIgnoreCase) then [Kind = "Option",      Body = Text.Trim(Text.Range(t, 7))]
            else if Text.StartsWith(t, "Explanation:", Comparer.OrdinalIgnoreCase) then [Kind = "Explanation", Body = Text.Trim(Text.Range(t, 12))]
            else if Text.StartsWith(t, "Image:",       Comparer.OrdinalIgnoreCase) then [Kind = "Image",       Body = Text.Trim(Text.Range(t, 6))]
            else [Kind = "Unknown", Body = t], type record),

    Split   = Table.ExpandRecordColumn(Classified, "Parts", {"Kind", "Body"}, {"Kind", "Body"}),
    Known   = Table.SelectRows(Split, each [Kind] <> "Unknown"),

    // ---------- 3. number the question blocks (the critical fix) ----------
    Indexed  = Table.AddIndexColumn(Known, "RowNo", 0, 1, Int64.Type),
    Anchored = Table.AddColumn(Indexed, "QAnchor", each if [Kind] = "Question" then [RowNo] else null, Int64.Type),
    Filled   = Table.FillDown(Anchored, {"QAnchor"}),
    Anchors  = List.Sort(List.Distinct(Table.Column(Filled, "QAnchor"))),
    Numbered = Table.AddColumn(Filled, "QuestionID", each List.PositionOf(Anchors, [QAnchor]) + 1, Int64.Type),
    Bodies   = Table.SelectRows(Numbered, each [QAnchor] <> null),

    // ---------- 4. build one row per question ----------
    Grouped = Table.Group(Bodies, {"QuestionID"}, {{"Group", each _}}),

    Built = Table.AddColumn(Grouped, "Built", each
        let
            G     = [Group],
            Pick  = (k as text) as list => Table.Column(Table.SelectRows(G, each [Kind] = k), "Body"),
            QList = Pick("Question"),
            OList = Pick("Option"),
            EList = Pick("Explanation"),
            IList = Pick("Image"),

            QText = if List.Count(QList) = 0 then "" else QList{0},
            Exp   = if List.Count(EList) = 0 then null else EList{0},
            Img   = if List.Count(IList) = 0 then null else IList{0},

            Flags   = List.Transform(OList, each Text.EndsWith(_, "*")),
            Clean   = List.Transform(OList, each Text.TrimEnd(Text.Trim(_), "*")),
            Kept    = List.FirstN(Clean, MaxOptions),

            // 1-based option numbers, capped at the 5 columns the template has
            Correct = List.Transform(
                          List.Select(List.Positions(Flags), each Flags{_} and _ < MaxOptions),
                          each _ + 1),

            Padded  = List.Transform({1..MaxOptions}, each if List.Count(Kept) >= _ then Kept{_ - 1} else null),
            QType   = if List.Count(Correct) > 1 then "Checkbox" else "Multiple Choice",

            Issues  = Text.Combine(List.RemoveNulls({
                          if List.Count(Correct) = 0            then "NO CORRECT ANSWER" else null,
                          if List.Count(OList) > MaxOptions      then "had " & Text.From(List.Count(OList)) & " options; extras dropped" else null,
                          if List.Count(OList) < 2               then "fewer than 2 options" else null,
                          if List.Count(IList) > 1               then "multiple Image: lines; first kept" else null,
                          if List.Count(EList) = 0               then "missing explanation" else null
                      }), "; "),

            Rec = Record.FromList(
                      {QText, QType} & Padded &
                      {Text.Combine(List.Transform(Correct, Text.From), ","), TimeSeconds, Img, Exp,
                       QText, List.Count(OList), List.Count(Correct), Issues},
                      {"Question Text", "Question Type", "Option 1", "Option 2", "Option 3",
                       "Option 4", "Option 5", "Correct Answer", "Time in seconds",
                       "Image Link", "Answer explanation",
                       "__QText", "__OptionCount", "__CorrectCount", "__Issues"})
        in
            Rec, type record),

    Expanded = Table.ExpandRecordColumn(Built, "Built",
        {"Question Text", "Question Type", "Option 1", "Option 2", "Option 3",
         "Option 4", "Option 5", "Correct Answer", "Time in seconds",
         "Image Link", "Answer explanation",
         "__QText", "__OptionCount", "__CorrectCount", "__Issues"}),

    Typed = Table.TransformColumnTypes(Expanded, {
        {"Correct Answer", type text},
        {"Time in seconds", Int64.Type},
        {"__OptionCount", Int64.Type},
        {"__CorrectCount", Int64.Type}
    }),

    Result = Table.Sort(Typed, {{"QuestionID", Order.Ascending}})
in
    Result
