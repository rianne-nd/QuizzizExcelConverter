# Quizizz Reviewer Toolkit

Turns AI-formatted quiz text into Quizizz-ready spreadsheets.

> **One launcher:** `Quiz Toolkit.cmd` now runs all five jobs — sort, convert, split,
> combine, and publish to Wayground — from a single menu. The older
> `Convert My Quizzes.cmd` and `Split My Quizzes.bat` still work, but you only need
> the toolkit.

> **Your quiz files live in `Quizzes\`.** Save each new `.txt` in its subject folder
> using the `Subject_WhatQuiz.txt` pattern — for example
> `Quizzes\DataComs\DataComs_Mod14.txt` or `Quizzes\AltOS\AltOS_Quiz2.txt`.
> Or drop it straight into the general `Quizzes\` folder and choose **Sort**:
> the toolkit reads the word before the first `_`, moves the file into
> `Quizzes\<Subject>\`, converts it, opens the output folder, and copies that
> folder path to the clipboard.

> **Maintainers / AI agents:** read `ARCHITECTURE.md` before changing code — it maps
> every component, file format, invariant, and extension point.

> **Just want to use it?** Read `START-HERE.md` instead — it has plain, numbered steps and
> assumes no Excel or PowerShell knowledge. This file is the technical explanation.

---

## 1. The bug that was silently corrupting your quizzes

Your old query built the question key like this:

```m
#"Added Custom" = Table.AddColumn(#"Removed Blank Rows", "RecordID",
    each if Text.StartsWith([Column1], "Question:") then 1 else null),
```

Every `Question:` line sets `RecordID = 1` — it **never increments**. Later steps grouped by

```
{"RecordID", "Question Text", "Answer explanation"}
```

so the grouping was really *just by question text*. When two questions shared the same text,
the two blocks were **merged into one**, producing 8 option rows for a 4-option question.

You can see the damage in `Quizizz_Mod14.xlsx`:

| Symptom | Example | Cause |
|---|---|---|
| Answer key out of range | `Correct Answer = 4,8` on a 4-option question | index 8 = option 4 of the merged 2nd copy |
| Phantom option | `Option 5` identical to `Option 1` | it is the 2nd copy's first option |
| Silently lost options | options 6–8 vanished | `List.Max([1])`…`List.Max([5])` drops them |

`Mod14Raw.txt` contains **20 intentionally repeated questions** — the same question placed
again later in the quiz as a recall item. That is a deliberate technique, not a mistake. The
bug meant those repeats were **merged into the original instead of appearing again**, so the
recall item never reached the quiz at all, and the two merged questions got a corrupted answer
key between them.

**After the fix**, the same file produces **87 separate rows** — the 67 distinct questions plus
the 20 repeats, each as its own valid question — with **0 out-of-range answers and 0 phantom
option 5s** (verified).

Repeats are therefore kept, never merged, and are only *counted* in the report. Pass
`-FlagDuplicates` if you ever want them listed as well — off by default, since flagging them
would be pure noise for your workflow.

### Four other bugs fixed at the same time

| # | Bug | Effect |
|---|---|---|
| 2 | `Table.ReplaceValue(..., "*", ReplaceText, ...)` removed **every** asterisk | destroyed emphasis like `*before*` inside option text |
| 3 | `FillUp` on "Answer explanation" across the whole file | a question with no explanation inherited the **next** question's explanation |
| 4 | `Image Link` was `each ""` | that column could never be filled |
| 5 | Options beyond 5 were dropped but the answer index was not | answers pointed at options that no longer existed |

---

## 2. The new workflow

```
source material
      |
      v
  Prompt.txt  ---->  AI  ---->  Quizzes\DataComs\DataComs_Mod14.txt
                                      |       (Question:/Image:/Option:/Explanation:)
                    +-----------------+-----------------+
                    |                                   |
                    v                                   v
        tools\Convert-Quizzes.ps1              Power Query (in Excel)
        one .xlsx per .txt                     ModuleRaw / QuizizzRawAll / QuizQA
```

**Power Query cannot write files** — M has no file-output function. That is why the bulk
"one Excel file per txt" job lives in a script, and Power Query handles the in-Excel path.

All five jobs are wrapped by one launcher:

| `Quiz Toolkit.cmd` option | What it runs |
|---|---|
| 1. Convert | `tools\Convert-Quizzes.ps1` — the same engine as the old `Convert My Quizzes.cmd` |
| 2. Sort & convert | built-in auto-sort: `Quizzes\<file>` → `Quizzes\<Subject>\`, then convert |
| 3. Split | built-in splitter — same rules as `tools\Splitter.py` |
| 4. Combine | built-in combiner — same rules as `tools\combiner.py` |
| 6. Upload existing XLSX | `wqu publish "<folder>" --assign Naic` / `--no-assign` (optional Python CLI) |

Splitting and combining no longer require Python. After a conversion the toolkit
also offers to publish the newly created `.xlsx` files with `wqu`: one class
prompt for the batch, then a Y/n for a single file or All / Select / None for
several. Uploading is skipped with a warning when `wqu` is not installed.

If you split a quiz that is still sitting in the general `Quizzes\` folder, the
toolkit files it into its subject folder first, so the original and every
`_ptN` part are created in `Quizzes\<Subject>\`, and the Excel output goes to
that subject's `_QuizizzOutput`.

**Auto-sort rules.** The prefix before the first `_` is matched case-insensitively
against the folders under `Quizzes\`. A new prefix creates a new subject folder, except
for module/checkpoint-style names (`Mod14`, `Checkpoint...`), which go to
`Quizzes\_Unsorted`. Files with no `Question:` lines go to `Needs-Fixing`. Same-name
collisions keep both files: identical content moves to `_backups\duplicates`, different
content gets a `(2)` suffix. After a successful sort and convert, the output folder
opens and its path is copied to the clipboard.

When you choose **Convert**, the toolkit lists the newest `.txt` files from `Quizzes\`,
newest first. Press Enter to use `[1]`, type another number to choose it, or type `B` to
open File Explorer. Type `F` to convert a whole folder instead of one file. If the chosen
file is sitting directly in `Quizzes\`, it is auto-sorted first.

---

## 3. Bulk convert: one `.xlsx` per `.txt`

```powershell
# the consolidated launcher (opens the menu)
& ".\Quiz Toolkit.cmd"

# or run the conversion directly, without the menu
& ".\Quiz Toolkit.cmd" -Action Convert -Path ".\Quizzes\DataComs"

# every .txt in the folder -> its own .xlsx
.\tools\Convert-Quizzes.ps1 -InputFolder ".\Quizzes\DataComs"

# a single file, into your existing output folder
.\tools\Convert-Quizzes.ps1 `
    -InputPath ".\Quizzes\DataComs\DataComs_Mod14.txt" `
    -OutputFolder ".\Quizzes\DataComs\_QuizizzOutput"

# include sub-folders
.\tools\Convert-Quizzes.ps1 -InputFolder ".\Quizzes" -Recurse

# sort every new .txt sitting directly in Quizzes\ (auto-detect subject),
# convert each one, open the output folders, and copy the last output path
& ".\Quiz Toolkit.cmd" -Action Sort

# same, but automation-friendly: no Explorer window, no clipboard change
& ".\Quiz Toolkit.cmd" -Action Sort -NoOpen -NoClipboard

# sort one specific dropped file
& ".\Quiz Toolkit.cmd" -Action Sort -Path ".\Quizzes\DataComs_Mod14.txt"
```

Naming: a trailing `Raw` / `_Raw` is stripped, then `Quizizz_` is prefixed.

| Input | Output |
|---|---|
| `DataComs_Mod14.txt` | `Quizizz_DataComs_Mod14.xlsx` |
| `AltOS_Quiz2.txt` | `Quizizz_AltOS_Quiz2.xlsx` |
| `Mod14Raw.txt` | `Quizizz_Mod14.xlsx` (the `Raw` suffix is still removed) |

Each workbook contains **exactly** the 11 Quizizz template columns as row 1 (bold, frozen),
so it can be imported directly or pasted into your template. A `_conversion-report.txt` is
written alongside the workbooks listing every question that needs attention.

Useful switches: `-TimeSeconds 30`, `-MaxOptions 5`, `-NamePrefix "Quizizz_"`,
`-Recurse`, `-OutputFolder <path>`.

The launcher accepts the same ideas without opening the menu:

```powershell
& ".\Quiz Toolkit.cmd" -Action Split -Path "C:\...\Quiz.txt" -QuestionsPerFile 20 -TailThreshold 15
& ".\Quiz Toolkit.cmd" -Action Combine -Path "C:\...\Quiz_pt1.txt"
```

---

## 4. Power Query setup

Import the four files from `powerquery/` as queries, **in this order**
(Data → Get Data → From Other Sources → Blank Query → Advanced Editor → paste → Done).
Then rename each query to the file name shown:

| File | Query name | Purpose |
|---|---|---|
| `fnParseQuizizzRaw.pq` | `fnParseQuizizzRaw` | the parser (function) — **create this first** |
| `ModuleRaw.pq` | `ModuleRaw` | one file → 11 template columns. Your old query's replacement |
| `QuizizzRawAll.pq` | `QuizizzRawAll` | a whole folder → one combined table with a `Source File` column |
| `QuizQA.pq` | `QuizQA` | only the questions that need a human look |

Edit the `FilePath` / `FolderPath` at the top of each to point at your data.

**`ModuleRaw` is the drop-in replacement for your `Module1Raw`.** Only `FilePath` changes
between modules.

### New: `Image:` support

The parser now reads an optional `Image:` line anywhere in a question block:

```
Question: Which of these is the largest planet?
Image: https://cdn.pixabay.com/photo/2014/09/08/09/24/solar-system-439046_960_720.png
Option: Earth
Option: Mars
Option: Jupiter*
Explanation: Jupiter is a gas giant.
```

* No `Image:` line → the cell is left **blank** (correct for image-less questions).
* The URL lands in the `Image Link` column automatically.
* More than one `Image:` line → the first is kept and `QuizQA` flags it.

---

## 5. What the prompt now asks the AI for

`Prompt.txt` gained rule 4 (image links). The important part is deliberate:

> **NEVER invent, guess, or fabricate a URL.** … If the source material clearly shows an
> image for that question but you cannot find its URL, write exactly `Image: PASTE_URL_HERE`.

Language models routinely produce plausible-looking but dead URLs, and Quizizz will happily
import a broken image. So the prompt only carries a link the source material actually
contained, and otherwise leaves you a searchable `PASTE_URL_HERE` placeholder.

To find the ones needing a link:

```powershell
Select-String -Path ".\Quizzes\DataComs\*.txt" -Pattern 'PASTE_URL_HERE'
```

It also now says: use `*` **only** as the answer marker, never inside text — which stops the
`*before*` ambiguity at the source.

Your original prompt is preserved as `_backups\Prompt_v1_backup.txt`.

---

## 6. Verification status — please read

| Component | Status |
|---|---|
| `Convert-Quizzes.ps1` | **Executed and verified.** Ran against the real `Mod14Raw.txt` (87 questions, 20 intentional repeats kept as separate rows) and a synthetic file covering images, zero-width characters, inline asterisks, missing explanations and 6-option questions. Also tested on Windows PowerShell 5.1 and through the `Convert My Quizzes.cmd` launcher. Generated `.xlsx` files were unzipped and re-read to confirm cell values. The reorganized `Quizzes\` folders were also converted end-to-end (16 workbooks). |
| `QuizToolkit.ps1` | **Executed and verified.** Split output compared byte-for-byte with `Splitter.py` on the real 74-question Checkpoint file (7 parts, 10/10/10/10/10/10/14) and on folder, preamble, and blank-line-fallback cases. Combine output compared byte-for-byte with `combiner.py` for both folder and single-file modes. Conversion and the menu flow were tested through `Quiz Toolkit.cmd`. The recent-file list was tested by adding a new quiz — it appeared as `[1]` and Enter converted it. The File Explorer and folder pickers were created and initialized in STA; the final click in the dialog is human-driven. Auto-sort was exercised end to end with temporary files: valid prefix (new subject folder + move + convert), malformed format (→ `Needs-Fixing`), module-style name (→ `_Unsorted`), same-name/different-content collision (`(2)` suffix), identical duplicate (→ `_backups\duplicates`), targeted file and folder modes, menu options 1 and 2, clipboard copy, and `-NoOpen`/`-NoClipboard`. |
| The bug fix | **Verified.** 87 rows, 0 out-of-range answers, 0 phantom `Option 5` (the old query produced 20). |
| `Prompt.txt` | Rewritten; not machine-testable. |
| The `.pq` M files — **syntax** | **Validated.** All four parse cleanly, using the Mashup parser that ships with Excel (see `tools\Test-PowerQuerySyntax.ps1`). The checker self-tests against deliberately broken M first, so a pass is meaningful. |
| The `.pq` M files — **execution** | **Not executed.** Evaluating M needs `Microsoft.Mashup.Container`, which talks over named pipes — blocked by this sandbox. That is why `Workbook.Queries.Add` fails even for `let x = 1 in x`. |

So: the M is syntactically clean and its *logic* is independently proven by the PowerShell
converter, which implements the same rules and produces the same result on the same input.
What remains unverified is M *runtime* behaviour (e.g. a wrong argument type that only shows
up when data flows through). Power Query reports that with a line number when you hit **Done**.

Run `.\tools\Test-PowerQuerySyntax.ps1` any time after editing a `.pq` file — it re-checks
syntax in about a second, without opening Excel. Paste `fnParseQuizizzRaw` first; it is the
only one with real logic.

---

## 7. Folder layout

The launcher stays in the main folder. Every program lives in `tools\`, and every
quiz source file lives under `Quizzes\`, grouped by subject.

```
START-HERE.md                    plain step-by-step guide - start here
README.md                        this file (user-level technical overview)
ARCHITECTURE.md                  maintainer/AI-agent map: components, formats, invariants
Quiz Toolkit.cmd                 double-click: sort + convert + split + combine menu
Convert My Quizzes.cmd           old convert-only launcher (still works)
Split My Quizzes.bat             old split-only launcher (still works)
Prompt.txt                       formatter prompt (existing questions)
GenerateQuizPrompt.txt           generator prompt (write questions from a lesson)

Quizzes\                         every quiz .txt, grouped by subject
  README.txt                     where to save + naming reminder
  _Unsorted\                     files without a usable subject prefix (manual fix)
  DataComs\                      DataComs_*.txt + _QuizizzOutput\
  AltOS\                         AltOS_*.txt + _QuizizzOutput\
  AppDev\                        AppDev_*.txt + _QuizizzOutput\
  OpSy\                          OpSy_*.txt + _QuizizzOutput\
  SocPro\                        SocPro_*.txt + _QuizizzOutput\
  WebDev\                        WebDev_*.txt + _QuizizzOutput\

Needs-Fixing\                    quizzes that need formatting fixes first
  OpSy_Mod3.txt                  wrong labels + no answer markers
  README.txt                     what is wrong and how to fix it

tools\                           every program, in one place
  QuizToolkit.ps1                engine behind Quiz Toolkit.cmd (all five jobs)
  Convert-Quizzes.ps1            conversion engine used by the toolkit
  Splitter.py                    old Python splitter (kept for compatibility)
  combiner.py                    old Python combiner (kept for compatibility)
  Test-PowerQuerySyntax.ps1      syntax-checks the .pq files without Excel
  test_splitter.py               tests for Splitter.py

powerquery\                      the four Power Query files
  fnParseQuizizzRaw.pq             parser function
  ModuleRaw.pq                     single file -> template columns
  QuizizzRawAll.pq                 folder -> combined table
  QuizQA.pq                        problem-question report

_backups\                        retired versions, kept just in case
  Before-Organization\           the old folders, split parts, and old outputs
  Before-Organization\MANIFEST.txt   original path -> new path for every move
  duplicates\                    identical copies auto-sort set aside
tests\                           test fixtures
```
