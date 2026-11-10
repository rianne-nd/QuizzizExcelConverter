# START HERE

You don't need to understand any of the technical stuff. Just follow the steps.

There are **two ways** to do this. You only need one.

* **Way 1 — the double-click way** (easiest, recommended). One program,
  **`Quiz Toolkit.cmd`**, sorts new files into the right subject folder, converts
  `.txt` files into Excel files, splits a long quiz into parts, combines parts
  back into one quiz, and can upload the finished Excel files to Wayground.
* **Way 2 — the Excel way.** Does it inside Excel, one file at a time. Only use this if
  you specifically want to work inside Excel.

Start with Way 1. The conversion part is three steps.

## Where your quizzes live

All quiz `.txt` files live under `Quizzes\`, one folder per subject:

```
Quizzes\
    DataComs\
    AltOS\
    AppDev\
    OpSy\
    SocPro\
    WebDev\
```

Name every new quiz `Subject_WhatQuiz.txt` — for example `DataComs_Mod14.txt`
or `AltOS_Quiz2.txt`.

**The easy way:** save or drop the new `.txt` file directly into the `Quizzes`
folder (the general folder). The toolkit reads the word before the first `_`
and files it into the right subject folder for you. See Step 3.

---

# WAY 1 — The double-click way

## The whole routine, in 3 steps

### Step 1 — Get the AI to format your text

1. Open the file `Prompt.txt` (it's in this folder).
2. Select everything in it (press `Ctrl` + `A`), then copy (press `Ctrl` + `C`).
3. Paste that into your AI chat. It will ask for your raw text.
4. Give it your raw notes, screenshots, reviewer, whatever.
5. The AI replies with a block of formatted text. Copy it.

### Step 2 — Save that text as a `.txt` file

1. Open Notepad (press the Windows key, type `notepad`, press Enter).
2. Paste the AI's text into Notepad (`Ctrl` + `V`).
3. Click **File → Save As**.
4. **Important:** at the bottom, change **"Save as type"** to **All Files (\*.\*)**.
   If you skip this, Notepad saves it as `DataComs_Mod14.txt.txt` and nothing will work.
5. Name it like `Subject_WhatQuiz.txt` — for example `DataComs_Mod14.txt`
   or `AltOS_Quiz2.txt`.
6. Save it directly into the `Quizzes\` folder — the general folder. You do
   **not** have to pick a subject folder; the toolkit sorts it for you. (You
   can still save straight into `Quizzes\DataComs\` and so on if you prefer.)
7. Do this for every `.txt` you want. You can make many of them.

### Step 3 — Sort and convert them

1. Double-click **`Quiz Toolkit.cmd`** in this folder.
2. A black window opens with a menu. Type `2` and press Enter (that's
   **Sort & convert new quizzes**).
3. The toolkit reads the word before the `_` in each new file name, moves the
   file into `Quizzes\<Subject>\`, converts it to Excel, opens the output
   folder, and copies that output folder path to your clipboard.
4. After the Excel files are made, it can offer to upload them to Wayground.
   See **Uploading to Wayground** below. You can always answer `N`.
5. If your file has no subject prefix (for example `Mod14Raw.txt`), it goes
   to `Quizzes\_Unsorted\` and the toolkit tells you. Rename it like
   `DataComs_Mod14.txt`, then run `2` again.
6. If the text is not in the `Question:` format, the file goes to
   `Needs-Fixing\` instead of being converted. Read
   `Needs-Fixing\README.txt` for the fix.

**Converting one file instead?** Choose `1` (Convert). It shows the newest
quiz files first:

* Press **Enter** to use the newest one.
* Type a number to pick another.
* Type `B` to choose with File Explorer.
* Type `F` to convert a whole folder.

If that newest file is sitting in the general `Quizzes\` folder, the toolkit
sorts it automatically before converting, then opens the output folder and
copies the path — same as option `2`.

That's it. Each Excel file is ready to import into Quizizz.

### Uploading to Wayground (after conversion)

If the `wqu` tool is installed, the toolkit offers to publish the files it
just converted. This is optional — answer `N` and the Excel files stay exactly
where they are.

1. It asks once for the whole batch:

   `Assign to class 'Naic'? (N for public 1-day deadline) [Y/n]`

   Press Enter or type `Y` to assign to the class **Naic**. Type `N` for a
   public assignment with the platform's default 1-day deadline.
2. If **one** Excel file was created, it asks:

   `Upload 'Quizizz_DataComs_Mod14.xlsx' to Wayground now? [Y/n]`

3. If **several** Excel files were created, it shows a numbered list and asks:

   `Upload these to Wayground? [A]ll, [S]elect specific, or [N]one:`

   * `A` uploads every file in the list.
   * `S` asks for numbers separated by commas, for example `1,3`.
   * `N` skips the upload.

4. The upload runs in the same black window, so you can watch `wqu` work and
   see the join links at the end. `wqu` also copies the links to the clipboard.

If `wqu` is not installed, the toolkit prints this warning and carries on with
the normal conversion:

> Python 'wqu' tool not found. Run pip install -e . in the wayground-quiz-uploader folder to enable auto-uploading.

To enable it, open a terminal in the separate `wayground-quiz-uploader` folder
(normally `C:\Users\Kath Puno\Documents\GithubProjects\wayground-quiz-uploader`),
run `pip install -e .`, then run `wqu login` once to sign in to Wayground.

> **If Windows shows a blue "Windows protected your PC" warning** when you double-click:
> click **More info**, then **Run anyway**. That's normal for a file like this.

> **Dragging is faster:** you can drag a file or folder from File Explorer and drop it directly
> onto `Quiz Toolkit.cmd`. The menu opens with it already selected.
> If you forget, type `B` at any file prompt to open File Explorer instead.

---

## What you end up with

Say `Quizzes\DataComs\` has these:

```
DataComs_Mod13.txt
DataComs_Mod14.txt
DataComs_Mod15.txt
```

After running it you get:

```
Quizzes\DataComs\_QuizizzOutput\
    Quizizz_DataComs_Mod13.xlsx
    Quizizz_DataComs_Mod14.xlsx
    Quizizz_DataComs_Mod15.xlsx
    _conversion-report.txt
```

One Excel file per text file. `Quizizz_` is added to the front. A trailing
`Raw` is still removed automatically if you happen to use it.

When you use option `2` (Sort & convert), the toolkit also opens the output
folder and copies its path to your clipboard, so you can paste it wherever
you need it. If `wqu` is installed, it then offers to upload the new files to
Wayground (see **Uploading to Wayground** above).

---

## Step 4 — Import into Quizizz (manual fallback)

If you let the toolkit upload to Wayground (see above), you can skip this step
and share the join link `wqu` printed. To import by hand:

Open one of the generated Excel files. Row 1 is the header, row 2 onwards is your quiz.
That is exactly the layout Quizizz expects.

* Try uploading the file to Quizizz directly.
* If Quizizz insists on its own template, open your template, click cell `A2`, then copy the
  generated file's rows (`Ctrl` + `A`, `Ctrl` + `C`) and paste them in (`Ctrl` + `V`).

---

## Read the report

Every time you run it, a file called **`_conversion-report.txt`** is written into the
`_QuizizzOutput` folder. Open it with Notepad. It looks like this:

```
[ OK ] DataComs_Mod14.txt  ->  Quizizz_DataComs_Mod14.xlsx   (87 questions, 20 recall repeats)
[WARN] DataComs_Mod15.txt  ->  Quizizz_DataComs_Mod15.xlsx   (40 questions)
         - Q12: NO CORRECT ANSWER marked - What does ARP resolve?
         - Q31: missing explanation - Which flag terminates a session?
```

* **`[ OK ]`** means nothing needs your attention.
* **`[WARN]`** means something is worth a look. The lines under it tell you which question
  number and what's wrong. Question numbers are counted from the top of that `.txt` file.
* **`recall repeats`** just tells you how many questions you deliberately repeated. That is
  normal and **not** a problem — see the note below.

The two warnings that actually matter:

| Warning | What to do |
|---|---|
| `NO CORRECT ANSWER marked` | You forgot the `*` at the end of the correct option in the `.txt`. Open the `.txt`, add the `*`, run it again. |
| `had 6 options; kept the first 5` | Quizizz allows only 5 options. If the correct answer was the 6th one, it got dropped — fix that question's `.txt`. |

---

## A note about your repeated questions

You repeat questions on purpose, as a recall item later in the quiz. That's a good idea, and
**this tool handles it correctly**: each repeat becomes its own separate question in the Excel
file, and the report just counts them.

⚠️ **Your old Power Query did not do this.** It merged any two questions that had the same
question text into a single broken question — one option list, two answer keys, and correct
answers pointing at options that didn't exist. So your recall repeats were probably never
actually making it into the quiz as separate questions. That is fixed now.

---

## If something goes wrong

| What you see | What it means | Fix |
|---|---|---|
| "Not found: ..." | The path you pasted is wrong or has a typo | Run it again and choose the option again, or paste the path carefully. You can copy a path from File Explorer's address bar. |
| No Excel files appear | There were no `.txt` files in that folder | Check you saved with "Save as type: All Files", and that the file ends in `.txt`. |
| `DataComs_Mod14.txt.txt` | Notepad added a second extension | Rename the file and delete the extra `.txt`. |
| A question is missing from the quiz | Its `.txt` block had no `Option:` lines | Check the report. |
| A file is sitting in `Needs-Fixing\` | It uses the wrong labels or has no answer markers | Read `Needs-Fixing\README.txt` for the exact fix. |
| Blue "Windows protected your PC" | Normal Windows warning | More info → Run anyway. |

If you get an error you don't understand, copy the text from the black window and send it to
me. That's the fastest way to get it fixed.

---

# WAY 2 — The Excel way (optional)

Use this only if you want to work inside Excel instead of generating files. You do this
**once** to set it up, then it's a refresh button after that.

You will paste four blocks of code into Excel. The code already exists as files in the
`powerquery` folder — you just copy and paste. **Do them in this exact order.**

### Step 1 — Create the first query

1. Open Excel with any workbook.
2. Click the **Data** tab at the top.
3. Click **Get Data** → **From Other Sources** → **Blank Query**.
   (The Power Query Editor window opens.)
4. In that window, click **Home** → **Advanced Editor**.
5. A box appears with some text in it. Click inside it, press `Ctrl` + `A` to select all,
   then press `Delete`.
6. Open the file `powerquery\fnParseQuizizzRaw.pq` in Notepad. `Ctrl` + `A`, `Ctrl` + `C`.
7. Go back to Excel and paste into the box (`Ctrl` + `V`). Click **Done**.
8. On the left there is a **Queries** panel listing your query (probably named `Query1`).
   **Right-click it → Rename → type exactly** `fnParseQuizizzRaw` and press Enter.

> This rename is essential. The other queries call it by that exact name.

### Step 2 — Create the second query

1. In the Power Query Editor: **Home** → **New Source** → **Other Sources** → **Blank Query**.
   (Or close the editor and repeat Data → Get Data → Blank Query.)
2. **Home** → **Advanced Editor** → select all → Delete.
3. Paste the contents of `powerquery\ModuleRaw.pq`. Click **Done**.
4. Right-click the new query → **Rename** → type exactly `ModuleRaw`.

### Step 3 — Point it at your file

1. Look at the **Applied Steps** list on the right of the Power Query Editor.
2. Find the step named **`Source`** and click the small **gear ⚙** icon next to it.
3. A file picker opens. Browse to your `.txt` file and choose it.
4. Click **Home** → **Close & Load**.

> If there is no gear on `Source`: click **Home** → **Advanced Editor** instead, and edit the
> line that starts with `FilePath =` so it holds your file's full path, in quotes. For example:
> `FilePath = "C:\Users\Kath Puno\Downloads\DataComs_Reviewer\Quizzes\DataComs\DataComs_Mod15.txt",`
> Then click **Done**.

Your quiz data appears in a new sheet, in the right columns, ready for Quizizz.

### Step 4 — Later, for a different module

1. Open the `ModuleRaw` query (double-click it in the Queries panel on the right of the sheet).
2. Click the gear ⚙ on the `Source` step again and pick a different `.txt`.
3. **Home** → **Close & Load**.

### Two extra queries, only if you want them

* `powerquery\QuizizzRawAll.pq` (name it `QuizizzRawAll`) — reads **every** `.txt` in a folder
  into one big table, with a `Source File` column so you can tell them apart.
* `powerquery\QuizQA.pq` (name it `QuizQA`) — shows **only** the questions that need fixing,
  across the whole folder. Handy before an exam.

You don't need either of these for Way 1 to work.

---

# Optional: splitting a long quiz into parts

If one `.txt` holds a lot of questions and you'd rather have them as several
shorter quizzes:

1. Double-click **`Quiz Toolkit.cmd`**.
2. Type `3` and press Enter (that's Split).
3. If nothing is selected yet, it shows your newest quiz files. Press Enter for the
   newest, type a number for another, type `B` to browse, or type `F` for a folder.
4. Press Enter twice to accept the defaults — 10 questions per file, and if the
   leftover is 15 or fewer, keep it together in one final file instead of
   making a tiny one.
5. Answer `y` at the end if you want it to convert the parts to Excel straight
   away. (It converts every `.txt` in that folder, including the original.)

It writes `YourQuiz_pt1.txt`, `_pt2.txt` and so on next to the original.

If the quiz is still in the general `Quizzes\` folder, the toolkit files it
into its subject folder first — so the parts land in `Quizzes\<Subject>\` and
the Excel output goes to that subject's `_QuizizzOutput`, not the general
folder.

Each part is a **whole number of questions**. A question is never cut in half,
and its explanation and any code block always travel with it.

---

# Optional: combining split parts back into one quiz

If you split a quiz and later want the single file back:

1. Double-click **`Quiz Toolkit.cmd`**.
2. Type `4` and press Enter (that's Combine).
3. If nothing is selected yet, drag a `_pt1.txt` file — or the folder holding
   the parts — onto the window and press Enter. You can also type `B` to browse
   or type `F` to pick a folder.

It writes `YourQuiz_combined.txt`. If you dropped a folder, the file goes into a
new **`Combined Quizzes`** sub-folder. If you dropped a single `_pt1.txt`, it is
written next to the parts.

---

# Optional: uploading an existing folder of Excel files (option 6)

Use this when you already have `.xlsx` files and do not want to convert
anything again.

1. Double-click **`Quiz Toolkit.cmd`**.
2. Type `6` and press Enter (that's **Upload existing XLSX to Wayground**).
3. The default folder is `Quizzes\DataComs\_QuizizzOutput`.
   * Press **Enter** to use that folder.
   * Type `B` to choose another folder with File Explorer.
   * Or drag a folder from File Explorer into the black window.
4. Answer the class question once:

   `Assign to class 'Naic'? (N for public 1-day deadline) [Y/n]`

   `Y` assigns to **Naic**; `N` makes it public with the platform's default
   1-day deadline.
5. `wqu` publishes every `.xlsx` directly inside that folder (sub-folders are
   not scanned) and prints the join links.

This uses the same `wqu` tool as the automatic offer. If it is missing, you
will see the install warning shown above.

---

# What's in this folder

The launcher stays in this main folder. Everything else that is a program lives
in `tools\`, and your quiz files live in `Quizzes\`.

| File / folder | What it's for |
|---|---|
| **`Quiz Toolkit.cmd`** | **Double-click this.** One menu for sorting, converting, splitting, combining, and uploading to Wayground. |
| **`Quizzes\`** | **The general folder. Drop new quiz `.txt` files here.** Option `2` sorts them into subjects. |
| `Quizzes\DataComs\`, `Quizzes\AltOS\`, etc. | Subject folders. Each gets its own `_QuizizzOutput\` when you convert. |
| `Quizzes\README.txt` | Short reminder of where to save and how to name new quizzes. |
| `Quizzes\_Unsorted\` | Files with no subject prefix (like `Mod14Raw.txt`) land here. Rename them and run option `2` again. |
| `Needs-Fixing\` | Quizzes that cannot be converted yet. Read the README inside. |
| `Convert My Quizzes.cmd` | Old convert-only launcher. Still works, but you no longer need it. |
| `Split My Quizzes.bat` | Old split-only launcher. Still works, but you no longer need it. |
| `Prompt.txt` | Copy-paste this into your AI when you already have the questions. |
| `GenerateQuizPrompt.txt` | Copy-paste this into your AI to write questions from a lesson. |
| `README.md` | The technical explanation, if you or someone helping you wants the details. |
| `ARCHITECTURE.md` | The maintainer/AI-agent map of the code, data formats, and extension points. |
| **`tools\`** | All the programs. You never need to open these by hand. |
| `tools\QuizToolkit.ps1` | The engine behind `Quiz Toolkit.cmd` — sorting, converting, splitting, combining, and the Wayground upload handoff. |
| `tools\Convert-Quizzes.ps1` | The conversion engine used by the toolkit. |
| `tools\Splitter.py` | Old Python splitter, kept for compatibility. |
| `tools\combiner.py` | Old Python combiner, kept for compatibility. |
| `tools\Test-PowerQuerySyntax.ps1` | Only needed if you edit the Power Query code by hand. |
| `powerquery\` | The four Power Query code files, for Way 2. |
| `_backups\` | Old versions and the previous folder layout, kept just in case. |
| `tests\` | Test files. You can ignore these. |
