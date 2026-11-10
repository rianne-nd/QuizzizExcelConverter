// ===========================================================================
//  QuizizzRawAll
//  Reads EVERY .txt in a folder and returns one combined table, with a
//  "Source File" column so you can filter to a single module.
//
//  This is the "multiple txt files" query. Power Query cannot WRITE files -
//  for one .xlsx per txt, use tools\Convert-Quizzes.ps1 instead.
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
        "Question Text", "Question Type", "Option 1", "Option 2", "Option 3",
        "Option 4", "Option 5", "Correct Answer", "Time in seconds",
        "Image Link", "Answer explanation",
        "__QText", "__OptionCount", "__CorrectCount", "__Issues"
    }),

    // Keep only the template columns + provenance. Add "__Issues" etc. back in
    // if you want the QA fields visible here.
    Result = Table.SelectColumns(Expanded, {
        "Source File",
        "Question Text", "Question Type", "Option 1", "Option 2", "Option 3",
        "Option 4", "Option 5", "Correct Answer", "Time in seconds",
        "Image Link", "Answer explanation"
    })
in
    Result
