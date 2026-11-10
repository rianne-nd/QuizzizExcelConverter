// ===========================================================================
//  ModuleRaw
//  Single-file, drop-in replacement for your old "Module1Raw" query.
//
//  To use a different module you have TWO options:
//    1. Click the gear icon next to the "Source" step in Applied Steps and
//       browse to a different .txt file. Easiest.
//    2. Edit the FilePath line below by hand.
//
//  Requires the fnParseQuizizzRaw function query to exist in this workbook
//  (Data > Get Data > Blank Query > Advanced Editor, then rename it).
//
//  Output: exactly the 11 columns of the Quizizz template, in template order,
//  so you can paste/load it straight into "Create a Quiz".
// ===========================================================================
let
    FilePath = "C:\Users\Kath Puno\Downloads\My Quizzes\Midterms\DataCom\Mod14Raw.txt",

    Source = File.Contents(FilePath),

    Parsed = fnParseQuizizzRaw(Source),

    Result = Table.SelectColumns(Parsed, {
        "Question Text",
        "Question Type",
        "Option 1",
        "Option 2",
        "Option 3",
        "Option 4",
        "Option 5",
        "Correct Answer",
        "Time in seconds",
        "Image Link",
        "Answer explanation"
    })
in
    Result
