# Architecture — Quiz Reviewer Toolkit

This document is the maintainer map for the whole project. It is written for
another programmer or an AI agent who needs to debug, extend, or continue the
work **without reading every script from the top**. Read this before changing
code.

---

## 1. What the system does

The toolkit turns AI-formatted quiz text into Quizizz-ready Excel workbooks,
publishes them to Wayground, and keeps the source `.txt` files organized by
subject.

Five jobs, one launcher:

| Job | Owner | Summary |
|---|---|---|
| Sort & convert | `tools\QuizToolkit.ps1` | Reads the word before the first `_` in a dropped file name, moves the file into `Quizzes\<Subject>\`, converts it, opens the output folder, copies the folder path to the clipboard. |
| Convert | `tools\Convert-Quizzes.ps1` | One `.xlsx` per `.txt`. Parses the raw format, runs QA checks, writes a minimal OOXML workbook. |
| Split | `tools\QuizToolkit.ps1` | Splits one quiz `.txt` into `_pt1.txt`, `_pt2.txt`, ... with N questions each. |
| Combine | `tools\QuizToolkit.ps1` | Merges `_ptN.txt` parts back into `_combined.txt`. |
| Publish to Wayground | `wqu` (external Python CLI) | After conversion the toolkit offers a one-time class choice (`--assign Naic` or `--no-assign`) and publishes the new workbooks. Menu option 6 uploads an existing folder on demand. |

Everything is Windows PowerShell 5.1 plus `cmd.exe`. No Excel and no
third-party PowerShell module is required. Conversion, sorting, splitting,
and combining need no Python. Publishing is optional and uses the separate
`wqu` tool from the `wayground-quiz-uploader` folder (`pip install -e .`);
when `wqu` is not installed the toolkit prints a warning and skips the upload.

**Non-goals:** the toolkit does not call an AI, does not convert PDFs or
screenshots, and does not fabricate quiz content. The AI formatting step
(`Prompt.txt` -> AI -> `.txt`) stays manual. A `.txt` that is not already in
the `Question:` format is parked in `Needs-Fixing\`, not guessed at.

---

## 2. Repository map

```
<main>\                                    workspace root (DataComs_Reviewer)
  Quiz Toolkit.cmd                         the launcher users double-click
  Convert My Quizzes.cmd                   old convert-only launcher (compatibility)
  Split My Quizzes.bat                     old Python split launcher (compatibility)
  START-HERE.md                            plain-language user guide
  README.md                                user-level technical overview
  ARCHITECTURE.md                          this file (maintainer map)
  Prompt.txt                               AI formatter prompt (existing questions)
  GenerateQuizPrompt.txt                   AI generator prompt (write from a lesson)
  QuizizzSampleSpreadsheetUpdated_v2.xlsx  reference template

  Quizzes\                                 live quiz sources, grouped by subject
    README.txt                             save/naming reminder for the user
    _Unsorted\                             files with no usable subject prefix
    DataComs\                              DataComs_*.txt + _QuizizzOutput\
    AltOS\  AppDev\  OpSy\  SocPro\  WebDev\
  Needs-Fixing\                            wrong-format files, awaiting a human
  _backups\                                retired files and reversible history
    Before-Organization\                   pre-sort folder layout + MANIFEST.txt
    duplicates\                            identical copies auto-sort set aside
  tests\                                   synthetic fixtures for the converter

  tools\
    QuizToolkit.ps1                        menu + auto-sort + split + combine
                                           + Wayground upload handoff
    Convert-Quizzes.ps1                    parser + QA + .xlsx writer
    Splitter.py                            legacy Python splitter (byte-compat reference)
    combiner.py                            legacy Python combiner (byte-compat reference)
    Test-PowerQuerySyntax.ps1              syntax-checks the .pq M files
    test_splitter.py                       tests for Splitter.py

  (separate checkout, not part of this repo)
  ..\..\Documents\GithubProjects\wayground-quiz-uploader\
    wqu\                                   optional Wayground publisher (Python)

  powerquery\                              Excel Power Query path (Way 2, optional)
    fnParseQuizizzRaw.pq                   parser function
    ModuleRaw.pq                           single file -> template columns
    QuizizzRawAll.pq                       folder -> combined table
    QuizQA.pq                              problem-question report
```

Generated at runtime:

```
Quizzes\<Subject>\_QuizizzOutput\
    Quizizz_<Name>.xlsx                    one workbook per source .txt
    _conversion-report.txt                 per-run QA report
```

---

## 3. Runtime entry points and parameters

### 3.1 `Quiz Toolkit.cmd`

```
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0tools\QuizToolkit.ps1" %*
```

* `-STA` matters: `System.Windows.Forms` file/folder pickers require a
  single-threaded apartment. `powershell.exe` is already STA in practice, but
  the flag makes it explicit and safe.
* `%*` forwards dragged paths and every switch.
* `cd /d "%~dp0"` makes relative paths (`.\Quizzes\...`) resolve from the root.
* The window pauses after a double-click or a non-zero exit.

### 3.2 `tools\QuizToolkit.ps1`

```
QuizToolkit.ps1
  [-Path <file-or-folder>]
  [-Action Menu|Convert|Split|Combine|Sort|Upload]   (default Menu)
  [-QuestionsPerFile <int>]                     (default 10)
  [-TailThreshold <int>]                        (default 15)
  [-Recurse]
  [-NoOpen]                                     (Sort only)
  [-NoClipboard]                                (Sort only)
```

* `Menu` — interactive menu. This is the normal path for the user.
* `Convert` — convert `-Path` (file or folder). Folder conversion asks about
  subfolders in the menu; the non-menu path uses `-Recurse`.
* `Split` — split `-Path`. Uses `-QuestionsPerFile` / `-TailThreshold`.
* `Combine` — merge `_ptN` parts under `-Path`.
* `Sort` — auto-sort. `-Path` is optional; with no path it scans the general
  `Quizzes\` folder (non-recursive).
* `Upload` — publish the folder or `.xlsx` file in `-Path` with
  `wqu publish ... --no-assign` (no prompts).
* `-NoOpen` / `-NoClipboard` — disable the automatic Explorer window and
  clipboard write for every auto-sort path: the `Sort` action, menu option 2,
  and the general-folder branch of menu option 1. They are for automation and
  tests.
* Menu option 6 asks for a folder (default
  `Quizzes\DataComs\_QuizizzOutput`), then one class choice, and runs
  `wqu publish "<folder>" --assign Naic` or `--no-assign`.

Script-scope state (set once, used by functions):

| Variable | Meaning |
|---|---|
| `$ScriptRoot` | `tools\` directory |
| `$WorkspaceRoot` | main folder (`tools\` parent) |
| `$QuizzesRoot` | `<main>\Quizzes` — the general folder |
| `$Converter` | path to `Convert-Quizzes.ps1` |
| `$Utf8NoBom` | `UTF8Encoding($false)` used for all text writes |

### 3.3 `tools\Convert-Quizzes.ps1`

```
Convert-Quizzes.ps1 [-InputFolder <path>]           # default: <main>\Quizzes\DataComs
Convert-Quizzes.ps1 -InputPath <file>               # single file
Convert-Quizzes.ps1 -InputFiles <file[]>            # explicit list, one shared report
  [-OutputFolder <path>]                            # default: <input>\..\_QuizizzOutput
  [-Recurse]                                        # include subfolders (folder mode)
  [-NamePrefix <string>]                            # default "Quizizz_"
  [-TimeSeconds <int>]                              # default 60
  [-MaxOptions <int>]                               # default 5 (Quizizz template limit)
  [-FlagDuplicates]                                 # list repeated question texts in the report
  [-StrictLines]                                    # drop unprefixed lines instead of joining them
```

`-InputFolder`, `-InputPath`, and `-InputFiles` are mutually exclusive parameter
sets. The last one exists so the toolkit can convert only the original + parts
from a split without touching other quizzes in the folder.

The default `InputFolder` is resolved **after** parameter binding, because
`$PSScriptRoot` is not available inside an advanced-function parameter default.

### 3.4 Legacy tools

* `tools\Splitter.py` — original splitter. Kept as a byte-for-byte reference.
  `QuizToolkit.ps1` reproduces its output exactly (verified).
* `tools\combiner.py` — original combiner. Same purpose.
* `Convert My Quizzes.cmd` — old convert-only launcher. It now defaults to
  `%~dp0Quizzes\DataComs` so it keeps working after the organization.
* `Split My Quizzes.bat` — calls `tools\Splitter.py` (Python required).

### 3.5 Power Query files (Way 2, optional)

`fnParseQuizizzRaw.pq` is the parser function; `ModuleRaw.pq` is the single-file
entry point; `QuizizzRawAll.pq` reads a folder; `QuizQA.pq` lists only questions
that need attention. `tools\Test-PowerQuerySyntax.ps1` parses the M syntax
without Excel. The M files are syntax-checked but **not executed** in this
environment (the Mashup container uses named pipes; see README section 6).

---

## 4. Data flows

### 4.1 Convert (`.txt` -> `.xlsx`)

```
txt file(s)
   |
   v
ConvertFrom-RawQuizText          (parse + QA, one question per block)
   |
   v
$parsed.Rows                    (ordered rows with the 11 Quizizz columns)
$parsed.Issues                  (warnings, one entry per problem)
   |
   v
Write-QuizizzXlsx               (minimal OOXML zip, no Excel required)
   |
   v
<OutputFolder>\Quizizz_<Name>.xlsx
<OutputFolder>\_conversion-report.txt
```

Naming (`Get-OutputBaseName`): strip a trailing `Raw` / `_Raw` / `_txt` token,
then prefix `NamePrefix` (default `Quizizz_`). With `-Recurse`, subfolder names
are flattened into the output name with `_`. Duplicate output names get `(2)`,
`(3)`, ... suffixes.

### 4.2 Auto-sort (`Quiz Toolkit.cmd` option 2, or `-Action Sort`)

```
.txt in the general Quizzes\ folder (or -Path target)
   |
   v
Get-QuizFormatInfo           count valid "Question:" and malformed "Question " lines
   |
   +-- 0 valid + malformed > 0  -> Needs-Fixing\        (wrong labels)
   +-- 0 valid                  -> Needs-Fixing\        (not a quiz)
   |
   v (valid)
Resolve-SubjectFolder        word before the first "_"
   |
   +-- no usable prefix        -> Quizzes\_Unsorted\
   +-- existing subject        -> use that folder (case-insensitive match)
   +-- genuinely new subject   -> create Quizzes\<Prefix>\
   |
   v
collision check on the target name
   +-- identical content       -> move to _backups\duplicates\
   +-- different content       -> keep both, target gets "(2)"
   |
   v
Move-Item                     source .txt -> Quizzes\<Subject>\
   |
   v
Invoke-Conversion             -> Quizzes\<Subject>\_QuizizzOutput\
   |
   v
open output folder(s) + copy last output folder path to clipboard
   |
   v
Invoke-SmartUpload            -> optional wqu publish (section 4.6)
```

The same pipeline runs inside option 1 (Convert) when the selected file is
sitting directly in the general `Quizzes\` folder. A file already inside a
subject folder is converted in place and not re-sorted.

### 4.3 Split

```
quiz .txt
   |
   +-- sitting in general Quizzes\ -> auto-sort (move only) to Quizzes\<Subject>\
   |
   v
Split-QuizText               split on lines starting with "Question:"
   |                         (fallback: blank-line blocks if none)
   v
Get-Chunks                   Size=N, TailThreshold=T
   |                         leftover <= T stays as one final chunk
   v
<name>_pt1.txt ... <name>_ptN.txt     (CRLF, UTF-8 no BOM)
```

The move-only sort is why a quiz dropped in the general folder is split inside
its subject folder: the original and all parts stay together, and the optional
post-split conversion writes to that subject's `_QuizizzOutput`.

The post-split "convert now?" offer passes the source file plus the newly
created parts to `Invoke-Conversion -Files`, which calls
`Convert-Quizzes.ps1 -InputFiles`. Unrelated `.txt` files in the same folder
are not reconverted.

### 4.4 Combine

```
folder or _pt1.txt
   |
   v
group files named  <base>_pt<number>.txt
   |
   v
concatenate in ascending part order
   |
   v
<base>_combined.txt
   folder input  -> <folder>\Combined Quizzes\
   single input  -> next to the parts
```

### 4.5 Recent-file suggestion

`Get-RecentQuizFiles` scans `Quizzes\` recursively, excludes build/test/backup
folders and non-quiz docs (see section 8), sorts by `LastWriteTime` descending,
and returns the newest 9. `Read-QuizTarget` displays them, accepts `Enter`
(newest), a number, `B` (File Explorer picker), `F` (folder picker), or a
dragged path.

### 4.6 Publish to Wayground (optional)

```
Invoke-Conversion / Invoke-AutoSort
   |
   v
array of full .xlsx paths (only files created or rewritten by this run)
   |
   v
Invoke-SmartUpload
   |
   +-- empty list       -> return quietly
   +-- wqu not found    -> warning, return
   |
   v
one Read-YesNo -> --assign Naic  |  --no-assign (public, 1-day deadline)
   |
   +-- 1 file  -> "Upload '<file>' to Wayground now?"
   +-- N files -> numbered list; A = all, S = pick numbers, N = none
   |
   v
wqu publish "<path1>" "<path2>" ... <assign flag>
   (synchronous; wqu keeps the live tree and prints the join links)
```

Menu option 6 asks for a folder (default
`Quizzes\DataComs\_QuizizzOutput`) and passes that folder to the same
`Invoke-WquPublish` helper. Every path is resolved to an absolute path with
`Resolve-Path` before it reaches `wqu`.

---

## 5. File formats

### 5.1 Raw quiz `.txt` (the input)

One field per line; a blank line is the conventional separator. Only
`Question:` is required to start a block.

```
Question: <text>
Image: <url>                     optional
Type: Fill-in-the-Blank          optional; only recognised type names count
Option: <text>
Option: <correct text>*          trailing "*" marks a correct answer
Answer: <accepted answer>        only for Fill-in-the-Blank
Explanation: <why>               optional but QA warns when missing
```

Rules implemented by `ConvertFrom-RawQuizText`:

* Prefixes are case-insensitive: `Question:`, `Option:`, `Answer:`,
  `Explanation:`, `Image:`, `Type:`.
* **Continuation lines**: an unrecognised line is appended to the previous
  text field (`Question`, `Explanation`, `Option`, or `Answer`) with a `\n`.
  This is how multi-line router configs and explanations survive.
* `Image:` and `Type:` deliberately do **not** change the "last field", so a
  config block after an `Image:` line still belongs to the question.
* The trailing `*` is read when the `Option:` line is first seen, **before**
  any continuation is appended. A later config line ending in `*` therefore
  cannot mark an option correct.
* `Type:` is only a directive when its value matches a known type
  (fill-in-the-blank / multiple choice / checkbox, with spelling variants).
  Prose like `Type: Extended ACLs are required...` falls through to
  continuation handling.
* Zero-width characters (`\u200B-\u200D\uFEFF`) and non-breaking spaces are
  stripped from every line.
* Every `Question:` starts a NEW block; two questions with identical text are
  never merged (the old Power Query bug).

### 5.2 Quizizz `.xlsx` (the output)

One sheet named `Create a Quiz`. Row 1 is bold and frozen. The 11 columns, in
exact order:

1. Question Text
2. Question Type
3. Option 1
4. Option 2
5. Option 3
6. Option 4
7. Option 5
8. Correct Answer
9. Time in seconds
10. Image Link
11. Answer explanation

Rules:

* `Correct Answer` is a comma-separated list of 1-based option numbers
  (`"2"`, `"1,3"`). One starred option -> **Multiple Choice**; two or more ->
  **Checkbox**.
* Fill-in-the-Blank (`Type:` line): accepted answers go in the Option columns
  and `Correct Answer` stays **blank** (what Wayground expects).
* Only the first `MaxOptions` (default 5) options are kept; extra options and
  any answer pointing beyond that are reported as issues.
* The writer is a plain zip: `[Content_Types].xml`, `_rels/.rels`,
  `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`,
  `xl/worksheets/sheet1.xml`. Strings are written as `inlineStr`, so no shared
  string table is needed.

### 5.3 Split and combined file names

* Split: `<originalBaseName>_pt<number>.txt` (e.g. `DataComs_Mod14_pt3.txt`).
* Combined: `<originalBaseName>_combined.txt`.
* `_ptN` files are excluded from recent-file suggestions and from auto-sort
  scans.

### 5.4 Subject folder convention

* The subject is the text before the **first** `_` in the file name:
  `DataComs_Mod14.txt` -> `DataComs`.
* Matching against existing folders is case-insensitive; the existing folder's
  spelling is reused (`datacoms_Mod1.txt` -> `Quizzes\DataComs\`).
* A truly new prefix creates a new folder, unless it looks like a module or
  checkpoint name. The "not a subject" test is:
  `^(?i)(mod|module|modules|checkpoint|checkoint|quiz|quizzes|exam|exams|lt|lq|raw|test|new|intro|introduction)\d*$`
  or an all-numeric prefix. Those go to `Quizzes\_Unsorted\` instead.
* Folder names are sanitized by removing `<>:"/\|?*` and trailing dots/spaces.
* Files with no `_` at all, or with the `_` at position 0, go to `_Unsorted`.

---

## 6. Function reference — `tools\QuizToolkit.ps1`

Functions are defined top-to-bottom; the script runs the Main section at the
bottom. `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'`
are active, so prefer `-LiteralPath` and explicit null checks.

### Console and path helpers

| Function | Purpose |
|---|---|
| `Clear-Screen` / `Write-Banner` | Screen reset and header. |
| `Read-MenuChoice -Prompt -Valid` | Validated single-key menu input. |
| `ConvertTo-CleanPath -Value` | Trims quotes, expands env vars and `~`, drops a trailing slash except drive/share roots. |
| `Resolve-ExistingPath -Value [-Label]` | Cleans + `Test-Path` + `Resolve-Path`; throws `Not found: ...`. |
| `Show-FilePicker [-Title] [-InitialDirectory]` | `OpenFileDialog` for `.txt`. |
| `Show-FolderPicker [-InitialDirectory] [-Description]` | `FolderBrowserDialog`; starts at the given folder, else `Quizzes\`. |
| `Read-Path -Prompt [-AllowEmpty]` | Path entry that also accepts `B` (file picker) and `F` (folder picker). |
| `Read-Int -Prompt -Default -Minimum` | Integer with a default. |
| `Read-YesNo -Prompt -Default` | Y/n with default. |
| `Test-WquInstalled` | `Get-Command wqu`; prints the install warning and returns false when absent. |
| `Read-AssignArguments` | One prompt; returns `--assign Naic` or `--no-assign`. |
| `Wait-Menu` | Pause until Enter. |

### Recent-file helpers

| Function | Purpose |
|---|---|
| `Get-RecentQuizFiles [-Count]` | Newest `.txt` under `Quizzes\`, filtered (section 8), default 9. |
| `Show-RecentQuizzes -Files` | Renders the numbered list with relative paths and timestamps. |
| `Read-QuizTarget [-Prompt]` | The complete "which quiz?" prompt: list, `Enter`, number, `B`, `F`, drag. Returns a full path. |

### Text helpers

| Function | Purpose |
|---|---|
| `Read-Utf8Text -FilePath` | `[IO.File]::ReadAllText(..., UTF8)`. |
| `Write-Utf8Text -FilePath -Text` | Normalizes to CRLF and writes UTF-8 **without BOM**. |

### Splitter

| Function | Purpose |
|---|---|
| `Split-QuizText -Text` | Splits on `Question:` lines; returns `Questions` + `Preamble`. |
| `Split-Into-Blocks -Text` | Fallback: blank-line separated blocks. |
| `Get-Chunks -Items -Size -TailThreshold` | Chunking with a tail threshold. |
| `Invoke-SplitFile -File -Size -TailThreshold` | One file -> `_ptN.txt` files; returns `{Success, Folder, Count, CreatedFiles}`. |
| `Invoke-Split -Target -Size -TailThreshold` | File or folder -> runs `Invoke-SplitFile` over it. A file in the general `Quizzes\` folder is first filed into its subject folder with `Invoke-AutoSort -SkipConversion`. Returns `{Success, Folder, Folders, SourceFiles, CreatedFiles, Files, SplitCount}`; `Files` is the source plus new parts, ready for the post-split conversion. |

### Combiner

| Function | Purpose |
|---|---|
| `Invoke-Combine -Target` | Groups `_ptN` files, writes `_combined.txt`; folder input goes to `Combined Quizzes\`. Returns `{Success, OutputFolder, Combined}`. |

### Converter wrapper

| Function | Purpose |
|---|---|
| `Get-ExcelFileSnapshot -Folder` | Records length + `LastWriteTimeUtc` for every `.xlsx` before a conversion. |
| `Get-NewExcelFiles -Folder -Before` | Returns the workbooks that are new or were rewritten, as full paths. |
| `Get-ConversionOutputFolder -Target` | The `_QuizizzOutput` folder for a file or folder target. |
| `Invoke-Conversion -Target [-RecurseSubfolders]` | Calls `Convert-Quizzes.ps1` with the right parameter set; returns an array of the full `.xlsx` paths created or rewritten by this run. |
| `Invoke-Conversion -Files <string[]>` | Converts exactly those `.txt` files (grouped by folder, each group gets its own `_QuizizzOutput` and report). Used by the post-split offer. |

### Wayground upload

| Function | Purpose |
|---|---|
| `Invoke-WquPublish -Paths -AssignArguments` | Resolves every path, prints the command, then runs `wqu publish ...` synchronously with the live tree in the current window. |
| `Read-UploadFolder [-DefaultFolder]` | Folder entry for menu option 6: Enter = default, `B` = folder picker, or drag a folder. |
| `Invoke-SmartUpload -Files` | The post-conversion offer: empty list returns; missing `wqu` warns; one class prompt for the batch; one-file Y/n or multi-file All/Select/None. |

### Auto-sort

| Function | Purpose |
|---|---|
| `Set-ClipboardText -Value` | `Set-Clipboard`, with a `clip.exe` fallback. Returns success. |
| `Get-AvailableFileName -Folder -BaseName -Extension` | Adds `(2)`, `(3)`, ... until the name is free. |
| `Test-IsGeneralFolderPath -FolderPath` | True when the folder is exactly `Quizzes\`. |
| `Get-QuizFormatInfo -FilePath` | Counts valid `Question:` lines and malformed `Question ` lines, after zero-width/NBSP cleanup. |
| `Get-CleanSubjectName -Value` | Strips invalid Windows filename characters. |
| `Resolve-SubjectFolder -FileName` | Prefix -> existing/new subject folder, or `$null` for no usable subject. |
| `Invoke-AutoSortCore [-Target] [-SkipConversion]` | The full sort pipeline; returns `{Success, Sorted, Skipped, NeedsFixing, Failed, SortedFiles, OutputFolders, CreatedFiles, CopiedPath}`. `-SkipConversion` moves the file into its subject folder without converting it. |
| `Invoke-AutoSort [-Target] [-SkipConversion] [-PassThruResult]` | Public wrapper; returns the created `.xlsx` paths, or the full result object with `-PassThruResult`. |

### Menu

| Function | Purpose |
|---|---|
| `Show-Menu -SelectedPath` | Renders the screen and option list. |
| `Read-ConvertTarget` | Thin wrapper over `Read-QuizTarget`. |
| `Start-Menu -InitialPath` | Main loop: options `1` Convert, `2` Sort & convert, `3` Split, `4` Combine, `5` Pick, `6` Upload existing XLSX, `0` Quit. |

---

## 7. Function reference — `tools\Convert-Quizzes.ps1`

The script is standalone: it can be run directly, or called by the toolkit.

| Function | Purpose |
|---|---|
| `ConvertFrom-RawQuizText -Path -MaxOptions -TimeSeconds -Columns [-FlagDuplicates] [-StrictLines]` | Parses one `.txt` into Quizizz rows and a list of issues. This is the heart of the system. |
| `Get-ColumnLetter -Number` | `1 -> A`, `27 -> AA`, for cell references. |
| `ConvertTo-XmlText -Value` | XML-escapes a cell value and removes control characters. |
| `Write-QuizizzXlsx -Rows -Columns -Path` | Builds the minimal OOXML zip described in 5.2. |
| `Get-OutputBaseName -File -Root -Prefix -Flatten` | Output naming (`Raw` stripping, `Quizizz_` prefix, subfolder flattening). |

Main flow (`Convert-Quizzes.ps1` bottom):

1. Resolve `InputFolder` / `InputPath` (default `Quizzes\DataComs` if neither).
2. Build the file list: single file, an explicit `-InputFiles` list, or
   `Get-ChildItem -Filter *.txt` plus `-Recurse` when requested. Skipped from
   the list: `~$*` Office temp files, `README*`, and anything inside
   `_QuizizzOutput\`.
3. For each file, parse, compute the output name, write the workbook, and
   append report lines.
4. Write `_conversion-report.txt` next to the workbooks.

### QA rules (all land in `_conversion-report.txt`)

| Rule | Meaning |
|---|---|
| `NO CORRECT ANSWER marked` | No option ends in `*`. |
| `had N options; kept the first 5` | More than `MaxOptions`; the dropped ones may include the answer. |
| `only N option(s)` | Fewer than two options. |
| `missing explanation` | No `Explanation:` line. |
| `says 'Choose two' but N answer(s) are starred` | Stem/key mismatch. |
| `ALL N options are marked correct` | Almost always a mistake. |
| `repeated option text` | Two identical options in one question. |
| `declared 'Multiple Choice' but N starred` / `declared 'Checkbox' but N starred` | `Type:` disagrees with the asterisks; asterisks win. |
| `Answer: lines but no 'Type: Fill-in-the-Blank'` | Typed answers ignored. |
| `Image: lines; kept the first` | More than one image URL. |
| `question text is N characters (long)` | Usually a pasted config block; Quizizz shows a small box. |
| `Fill-in-the-Blank with NO accepted answers` | FIB with no `Answer:` / `Option:` values. |
| `Fill-in-the-Blank has BOTH Answer: and Option: lines` | Answers were used; options ignored. |
| `N line(s) ignored` | Only with `-StrictLines`; unprefixed lines were dropped instead of joined. |

Repeated question **texts** are counted (`N recall repeats`) but not reported
as problems unless `-FlagDuplicates` is used. Each repeat stays a separate row.

---

## 8. Invariants and design decisions

These are load-bearing. Changing one usually breaks a documented behavior or a
byte-compatibility test.

1. **Repeats are never merged.** Every `Question:` line starts a new row, even
   when the text is identical to an earlier question. The old Power Query
   grouped by question text and corrupted answer keys; that must not come back.
2. **The asterisk is read before continuations.** `Option:` parsing removes the
   trailing `*` immediately; a later config line ending in `*` cannot mark an
   option correct.
3. **`Image:` and `Type:` do not change `LastField`.** Continuation lines after
   them still belong to the previous Question/Option/Explanation.
4. **A `Type:` line is a directive only for known values.** Prose starting with
   `Type:` falls through to continuation handling.
5. **Split/combine write CRLF with UTF-8 no BOM**, matching the old Python tools
   byte-for-byte.
6. **The converter skips `README*` and `_QuizizzOutput\`.** This makes
   `-Recurse` safe on the `Quizzes\` tree; `_conversion-report.txt` is never
   re-converted.
7. **Recent-file scanning excludes** `_QuizizzOutput`, `_backups`, `tests`,
   `tools`, `powerquery`, `Needs-Fixing`, `_Unsorted`, `_ptN` files, and
   `README` / `Prompt` / `GenerateQuizPrompt` files. Only real quiz sources are
   suggested.
8. **Auto-sort only trusts the file name, not the file contents, for the
   subject.** Content is used only to decide whether the file is convertible
   (`Question:` lines).
9. **Auto-sort collision policy keeps data.** Identical content -> duplicates
   backup; different content -> both files kept with a `(2)` suffix. Nothing is
   silently overwritten.
10. **The subject folder name wins on case.** Existing folders are matched
    case-insensitively and reused, so the folder layout stays predictable.
11. **`Needs-Fixing` and `_Unsorted` are manual parking lots.** Auto-sort leaves
    files already there alone (except reporting them), so repeated runs are
    idempotent.
12. **Clipboard and Explorer opening are side effects, not requirements.**
    `-NoOpen` / `-NoClipboard` keep the sort logic usable in tests and scripts.
13. **No Excel dependency.** The `.xlsx` writer emits the minimum OOXML parts.
14. **Windows-only assumptions.** `System.Windows.Forms`, `Set-Clipboard` /
    `clip.exe`, `Start-Process`, `cmd.exe`, and `-STA` are all Windows features.
15. **Post-split conversion is scoped to the split family.** The offer after a
    split converts only the source file and its new `_ptN` parts (via
    `-InputFiles`), never every `.txt` in the folder.

---

## 9. Configuration and constants

| Item | Value | Where |
|---|---|---|
| Workspace root | `tools\` parent | `$WorkspaceRoot` in `QuizToolkit.ps1` |
| General folder | `<main>\Quizzes` | `$QuizzesRoot` |
| Unsorted folder | `<main>\Quizzes\_Unsorted` | `Invoke-AutoSort` |
| Needs-Fixing root | `<main>\Needs-Fixing` | `Invoke-AutoSort` |
| Duplicates backup | `<main>\_backups\duplicates` | `Invoke-AutoSort` |
| Recent files shown | 9 | `Read-QuizTarget` |
| Split size | 10 | `-QuestionsPerFile` |
| Tail threshold | 15 | `-TailThreshold` |
| Quizizz columns | 11, exact order | `$QuizizzColumns` |
| Time column | 60 seconds | `-TimeSeconds` |
| Max options | 5 | `-MaxOptions` |
| Output prefix | `Quizizz_` | `-NamePrefix` |
| Fill-in-the-Blank types | fill-in-the-blank variants, fib, blank | `$FibPattern` |
| Multiple Choice / Checkbox | multiple choice, mc / checkbox, cb | `$McPattern`, `$CbPattern` |
| Non-subject prefixes | mod, module, checkpoint, checkoint, quiz, exam, lt, lq, raw, test, new, intro, numeric | `Resolve-SubjectFolder` |

---

## 10. How to extend

### Add a new menu action

1. Add the key to the `Read-MenuChoice -Valid` list in `Start-Menu`.
2. Add a `switch` case for it. Keep the `try/catch` + `Wait-Menu` pattern.
3. Add the line to `Show-Menu`.
4. If it needs to run without the menu, add the name to the `Action`
   `ValidateSet` and a case in the Main section, then document it in
   `START-HERE.md` and `README.md`.
5. If it writes files, respect `$Utf8NoBom` / `Write-Utf8Text` for CRLF output.

### Add a new parser directive

1. Add a `StartsWith(...)` branch in `ConvertFrom-RawQuizText` (case-insensitive).
2. Decide whether it changes `$cur.LastField`. Default: it should **not**, so
   continuation lines keep flowing into the previous text field.
3. Add QA rules and report lines near the other checks.
4. Update `Prompt.txt` so the AI emits the new directive.
5. Add a fixture under `tests\` and run the converter on it.
6. If Power Query should match, update `fnParseQuizizzRaw.pq`, then run
   `.\tools\Test-PowerQuerySyntax.ps1`.

### Change output naming

Edit `Get-OutputBaseName` in `Convert-Quizzes.ps1`. It currently strips
`([_\-\s]*raw)$` and `([_\-\s]*txt)$`, then prefixes `Quizizz_`; with
`-Flatten` it prepends the relative subfolder path with `_`. Keep the
duplicate-name loop intact.

### Add a new subject

* Easiest: drop `NewSubject_Quiz1.txt` into `Quizzes\` and run option `2`. The
  prefix is not in the non-subject list, so `Quizzes\NewSubject\` is created.
* If the name would collide with a real subject, add the prefix to the
  non-subject regex only when you want it parked in `_Unsorted` instead.

### Add or change the auto-sort heuristic

Everything lives in `Resolve-SubjectFolder` and `Invoke-AutoSort`. Watch three
things: (a) keep the case-insensitive existing-folder match, (b) keep the
collision policy non-destructive, (c) do not move files already in
`Needs-Fixing` / `_Unsorted`.

### Add a test recipe

Use a temp directory (for example
`C:\Users\<you>\AppData\Local\Temp\opencode\...`), create `.txt` fixtures,
call `Quiz Toolkit.cmd -Action Sort -Path ... -NoOpen -NoClipboard`, and assert
on the resulting folder layout. Clean up afterward. Converter fixtures live
under `tests\` (see `tests\fib`, `tests\continuation`, `tests\raw`).

---

## 11. Testing and verification

### Syntax check (fast, no side effects)

```powershell
$paths = @('.\tools\QuizToolkit.ps1', '.\tools\Convert-Quizzes.ps1')
foreach ($path in $paths) {
    $errors = $null; $tokens = $null
    [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count) { $errors | ForEach-Object { "Line $($_.Extent.StartLineNumber): $($_.Message)" } }
    else { "OK $path" }
}
```

### Converter smoke test

```powershell
.\tools\Convert-Quizzes.ps1 -InputFolder ".\Quizzes\DataComs"
```

Expect one workbook per `.txt`. The current organized tree converts to
**16 workbooks / 737 questions**.

### Split/combine byte-compatibility

`QuizToolkit.ps1` is verified byte-for-byte against `Splitter.py` on the real
74-question Checkpoint file (7 parts: 10/10/10/10/10/10/14) and against
`combiner.py` for both folder and single-file modes. Re-run the comparison
after touching `Split-QuizText`, `Get-Chunks`, `Invoke-SplitFile`,
`Invoke-Combine`, or `Write-Utf8Text`.

### Auto-sort test recipe

Create a valid `ZZTest_Quiz1.txt` in `Quizzes\`, then:

```powershell
& ".\Quiz Toolkit.cmd" -Action Sort -Path ".\Quizzes\ZZTest_Quiz1.txt" -NoOpen -NoClipboard
```

Check: the file moved to `Quizzes\ZZTest\`, `Quizizz_ZZTest_Quiz1.xlsx` exists
in `Quizzes\ZZTest\_QuizizzOutput\`, and the source is gone. Also test:
malformed labels (-> `Needs-Fixing`), `Mod99Raw.txt` (-> `_Unsorted`),
same-name/different-content (-> `(2)` suffix), identical copy
(-> `_backups\duplicates`), and clipboard (drop `-NoClipboard`, then
`Get-Clipboard`). Delete the temporary folders afterward.

### Verified status (2026-11)

* Converter: executed on the real files and synthetic edge cases; workbooks
  re-read from the zip.
* Split/combine: byte-for-byte equal to the Python tools.
* Auto-sort: exercised end to end (valid, malformed, unsorted, collisions,
  duplicates, targeted file/folder, menu options 1 and 2, clipboard,
  `-NoOpen`/`-NoClipboard`).
* Recent-file list: verified with a newly added file.
* File/folder pickers: created and initialized in STA; the final click is
  human-driven.
* Power Query M: syntax validated; execution not possible in this sandbox.

---

## 12. Known limitations / future work

* **No AI formatting step.** A wrong-format `.txt` goes to `Needs-Fixing\`;
  the toolkit cannot invent `Question:` labels or answer keys.
* **Prefix heuristic can be wrong for unusual names.** A file like
  `Mod14Raw.txt` has no subject, so it goes to `_Unsorted`. A new class must
  use a prefix that is not in the non-subject regex.
* **Windows-only.** Forms, clipboard, `clip.exe`, `Start-Process`, and `-STA`
  are Windows assumptions.
* **No undo besides backups.** Moves are reversible through
  `_backups\Before-Organization\MANIFEST.txt` and `_backups\duplicates\`, but
  there is no automatic "undo last sort" command.
* **Folder conversion flattens names.** `-Recurse` puts every workbook in one
  `_QuizizzOutput\` with `_`-joined subfolder names; it does not mirror the
  tree.
* **Power Query runtime is untested** in this environment.
* **The external `C:\Users\...\Downloads\My Quizzes\` collection is not
  managed** by this workspace; it was left untouched during the 2026-11
  organization.
* **`OpSy_Mod3.txt` is parked** in `Needs-Fixing\`: wrong labels and no `*`
  answer markers.

---

## 13. Maintainer troubleshooting

| Symptom | Likely cause | Check |
|---|---|---|
| "The Quizzes folder was not found" | `Quizzes\` missing or renamed | `$QuizzesRoot` in `QuizToolkit.ps1`; re-create the folder. |
| Sort created an unwanted subject folder | Prefix was not in the non-subject regex | `Resolve-SubjectFolder`; rename the file or move the folder. |
| File landed in `_Unsorted` | No usable prefix before `_` | Rename to `Subject_WhatQuiz.txt`, run option 2 again. |
| File landed in `Needs-Fixing` | No `Question:` lines | `Get-QuizFormatInfo`; fix labels and answer markers. |
| Clipboard did not update | `Set-Clipboard` and `clip.exe` both failed | Run in an interactive desktop session; check `-NoClipboard`. |
| Output folder did not open | `-NoOpen` set, or `Start-Process` failed silently | Remove `-NoOpen`; check the reported output path. |
| Report line says `6 options; kept the first 5` | Quizizz template limit | Fix the source question. |
| Workbook looks empty | Parser found no `Question:` lines | Run `Get-QuizFormatInfo` or check `_conversion-report.txt`. |
| Split output differs from `Splitter.py` | CRLF or question detection changed | Compare with `Write-Utf8Text` and `Split-QuizText`. |
| Menu key numbers feel wrong | Menu changed in 2026-11: Sort is `2`, Split `3`, Combine `4`, Pick `5` | `Show-Menu` / `Start-Menu`. |

When in doubt, run the syntax check first, then the converter smoke test, then
the auto-sort recipe with `-NoOpen -NoClipboard`. Those three cover the most
likely regressions without touching the real quiz tree.








