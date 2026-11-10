// ===========================================================================
//  QuizQA
//  Quality-check view over every raw .txt in a folder.
//  Returns only the questions that need a human look, so you can fix the raw
//  text BEFORE importing to Quizizz.
//
//  Flags:
//    * no correct answer marked      -> Quizizz will reject / mark it wrong
//    * more than 5 options           -> extras dropped, answer key capped
//    * fewer than 2 options
//    * missing explanation
//    * several Image: lines (only the first is used)
//
//  NOTE: repeated question text is NOT flagged. Repeating a question later in
//  the quiz is intentional (recall / spaced repetition), and each repeat is
//  kept as its own separate question row.
//
//  Requires fnParseQuizizzRaw in this workbook.
// ===========================================================================
let
    FolderPath = "C:\Users\Kath Puno\Downloads\My Quizzes\Midterms\DataCom",

    Files = Folder.Files(FolderPath),

    TxtOnly = Table.SelectRows(Files, each
        Text.Lower([Extension]) = ".txt"
        and not Text.StartsWith([Name], "~$")
        and not Text.StartsWith([Name], "_")),

    WithRows = Table.AddColumn(
        TxtOnly,
        "__Rows",
        each fnParseQuizizzRaw([Content]),
        type table
    ),

    WithSource = Table.AddColumn(WithRows, "Source File", each [Name], type text),

    Expanded = Table.ExpandTableColumn(WithSource, "__Rows", {
        "QuestionID", "Question Text", "Question Type", "Correct Answer",
        "Image Link", "__OptionCount", "__CorrectCount", "__Issues"
    }),

    ProblemRows = Table.SelectRows(Expanded, each
        (try [__Issues] <> "" otherwise false)
        or [Correct Answer] = ""
        or [Correct Answer] = null),

    Result = Table.SelectColumns(ProblemRows, {
        "Source File", "QuestionID", "Question Text", "Question Type",
        "Correct Answer", "__OptionCount", "__CorrectCount", "__Issues"
    }),

    Sorted = Table.Sort(Result, {{"Source File", Order.Ascending}, {"QuestionID", Order.Ascending}})
in
    Sorted
