#Requires -Version 5.1
<#
.SYNOPSIS
    Converts Quizizz "raw" .txt files into one Quizizz-ready .xlsx per .txt file.

.DESCRIPTION
    Raw format (one item per line, blocks separated by a blank line):

        Question: <question text>
        Image: <url>                <- OPTIONAL, only when the question has an image
        Option: <choice>
        Option: <correct choice>*   <- a trailing * marks a correct answer
        Explanation: <why>

    Question types are worked out automatically:
        no asterisk / one asterisk  -> "Multiple Choice"
        two or more asterisks       -> "Checkbox"

    To force a Fill-in-the-Blank question, add a Type: line. The accepted typed
    answers then go in Answer: lines (or Option: lines if you prefer), and
    "Correct Answer" is left blank - which is what Wayground expects:

        Question: Name any one of the two moons of Mars
        Type: Fill-in-the-Blank
        Answer: Phobos
        Answer: Deimos
        Explanation: Mars has two moons, Phobos and Deimos.

    Fill-in-the-Blank is NOT detected automatically, because a single accepted
    answer is indistinguishable from a normal one-answer Multiple Choice
    question. It only happens when you ask for it with a Type: line.

    CONTINUATION LINES
    A line with no recognised prefix is joined onto the field above it, instead
    of being thrown away. Real exam dumps write explanations and router config
    blocks across several lines:

        Explanation: Topic 3.10.2
        When data is encrypted, it is scrambled to keep the data private...

        Question: Consider the following access list.
        access-list 100 permit ip host 192.168.10.1 any
        access-list 100 deny icmp 192.168.10.0 0.0.255 any echo
        access-list 100 permit ip any any
        Which two actions are taken if the access list is placed inbound...?

    Both of those now survive, with the line breaks kept inside the cell.

    Three rules keep this safe:
      * A continuation joins the last Question / Option / Answer / Explanation.
        It never joins an Image: or Type: line, so a config block after an
        Image: line still belongs to the question.
      * The correct-answer asterisk is read when the Option: line is first seen,
        before any continuation is appended. A config line ending in "*" (common
        in "show ip route" output) therefore cannot mark an option correct.
      * Type: is only treated as a directive when its value is a real question
        type. Source dumps contain prose lines like
        "Type: Extended ACLs are required because ...", which are explanation
        text and are joined as such.

    Use -StrictLines to go back to dropping unprefixed lines and reporting them.

    Output: one .xlsx per input .txt containing exactly the 11 columns the Quizizz
    spreadsheet template expects, plus a plain-text conversion report.

    Repeated questions are KEPT as separate questions (this is how you add a
    recall item later in the quiz) and are only counted, never reported as errors.
    The old Power Query used to merge two questions that shared the same text,
    which corrupted the answer key of both.

    Power Query cannot write files, which is why this conversion lives in a script.

.PARAMETER InputFolder
    Folder containing the raw .txt files.
    Defaults to Quizzes\DataComs in the main folder.

.PARAMETER InputPath
    Convert a single .txt file instead of a whole folder.

.PARAMETER InputFiles
    Convert exactly these .txt files in one run (one shared report). Used by
    the toolkit after a split so unrelated quizzes in the folder are left
    alone. All files should live in the same folder; if they do not, the
    output folder defaults to the first file's folder.

.PARAMETER OutputFolder
    Where the .xlsx files are written.
    Defaults to "<input folder>\_QuizizzOutput".

.PARAMETER Recurse
    Also convert .txt files inside sub-folders. Sub-folder names are flattened
    into the output file name with '_'.

.PARAMETER NamePrefix
    Prefix for generated file names. Default 'Quizizz_'.
    A trailing "Raw"/"_Raw" is stripped first: Mod14Raw.txt -> Quizizz_Mod14.xlsx

.PARAMETER TimeSeconds
    Value written into the "Time in seconds" column. Default 60.

.PARAMETER MaxOptions
    Maximum options kept per question. Default 5 (the Quizizz template limit).
    Extra options are dropped and reported.

.PARAMETER FlagDuplicates
    Also list repeated question texts in the report. Off by default, because
    repeats are normally intentional recall items.

.EXAMPLE
    .\Convert-Quizzes.ps1
    Convert every .txt in the default Quizzes\DataComs folder.

.EXAMPLE
    .\Convert-Quizzes.ps1 -InputFolder ".\Quizzes\AltOS"

.EXAMPLE
    .\Convert-Quizzes.ps1 -InputPath ".\Quizzes\DataComs\DataComs_Mod14.txt" -OutputFolder ".\Quizzes\DataComs\_QuizizzOutput"
#>
[CmdletBinding(DefaultParameterSetName = 'Folder')]
param(
    [Parameter(ParameterSetName = 'Folder', Position = 0)]
    [string]$InputFolder,

    [Parameter(ParameterSetName = 'File', Mandatory = $true)]
    [string]$InputPath,

    [Parameter(ParameterSetName = 'Files', Mandatory = $true)]
    [string[]]$InputFiles,

    [string]$OutputFolder,

    [switch]$Recurse,

    [string]$NamePrefix = 'Quizizz_',

    [int]$TimeSeconds = 60,

    [int]$MaxOptions = 5,

    [switch]$FlagDuplicates,

    [switch]$StrictLines
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The default is resolved here rather than in the param block: $PSScriptRoot is
# not available while PowerShell is still evaluating parameter defaults.
if ([string]::IsNullOrWhiteSpace($InputFolder)) {
    $InputFolder = Join-Path (Split-Path -Parent $PSScriptRoot) 'Quizzes\DataComs'
}

# The Quizizz spreadsheet template header, in exact order.
$QuizizzColumns = @(
    'Question Text'
    'Question Type'
    'Option 1'
    'Option 2'
    'Option 3'
    'Option 4'
    'Option 5'
    'Correct Answer'
    'Time in seconds'
    'Image Link'
    'Answer explanation'
)

# ---------------------------------------------------------------------------
#  Parsing
# ---------------------------------------------------------------------------

function ConvertFrom-RawQuizText {
    <#
        Parses one raw .txt file into Quizizz rows.
        Every "Question:" line starts a NEW block - questions are never merged,
        even when two questions share the same text.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxOptions = 5,
        [int]$TimeSeconds = 60,
        [string[]]$Columns,
        [switch]$FlagDuplicates,
        [switch]$StrictLines
    )

    # A line with no recognised prefix is a CONTINUATION of the last text field,
    # not junk. The source material writes explanations and router config blocks
    # across several lines, so dropping them loses real content.
    # -StrictLines restores the old behaviour (drop them and count them).
    # One definition, used both to recognise a Type: line and to decide the row.
    # Keeping them separate earlier let "Fill-in-the-Blanks" (plural) through as
    # prose instead of a directive.
    $FibPattern = '^(?:fill[\s\-]*in[\s\-]*the[\s\-]*blanks?|fill[\s\-]*in[\s\-]*blanks?|fillintheblanks?|fib|blank)$'
    $McPattern  = '^(?:multiple[\s\-]*choice|mc)$'
    $CbPattern  = '^(?:checkbox|check[\s\-]*box|cb)$'
    $TypeNames  = '^(?:fill[\s\-]*in[\s\-]*the[\s\-]*blanks?|fill[\s\-]*in[\s\-]*blanks?|fillintheblanks?|fib|blank|multiple[\s\-]*choice|mc|checkbox|check[\s\-]*box|cb)$'

    $issues    = New-Object System.Collections.Generic.List[object]
    $questions = New-Object System.Collections.Generic.List[object]
    $cur       = $null
    $lineNo    = 0
    $unknown   = 0

    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path)) {
        $lineNo++
        if ([string]::IsNullOrEmpty($rawLine)) { continue }

        # Remove zero-width characters (they sneak in from PDF/Word copy-paste),
        # normalise non-breaking spaces, then trim.
        $line = $rawLine -replace '[\u200B-\u200D\uFEFF]', ''
        $line = ($line -replace '\u00A0', ' ').Trim()
        if ($line -eq '') { continue }

        $kind = $null
        $body = ''

        if     ($line.StartsWith('Question:',    [StringComparison]::OrdinalIgnoreCase)) { $kind = 'Question';    $body = $line.Substring(9).Trim() }
        elseif ($line.StartsWith('Option:',      [StringComparison]::OrdinalIgnoreCase)) { $kind = 'Option';      $body = $line.Substring(7).Trim() }
        elseif ($line.StartsWith('Answer:',      [StringComparison]::OrdinalIgnoreCase)) { $kind = 'Answer';      $body = $line.Substring(7).Trim() }
        elseif ($line.StartsWith('Explanation:', [StringComparison]::OrdinalIgnoreCase)) { $kind = 'Explanation'; $body = $line.Substring(12).Trim() }
        elseif ($line.StartsWith('Image:',       [StringComparison]::OrdinalIgnoreCase)) { $kind = 'Image';       $body = $line.Substring(6).Trim() }
        elseif ($line.StartsWith('Type:',        [StringComparison]::OrdinalIgnoreCase)) {
            # Only a RECOGNISED type name is treated as a directive. Source dumps
            # contain prose lines like "Type: Extended ACLs are required ..."
            # which are explanation text, not a question type.
            $candidate = $line.Substring(5).Trim()
            if ($candidate -match $TypeNames) { $kind = 'Type'; $body = $candidate }
        }

        if (-not $kind) {
            # --- continuation line ---
            if ($StrictLines -or -not $cur -or -not $cur.LastField) {
                $unknown++
                continue
            }
            $cur.Continuations++
            switch ($cur.LastField) {
                'Question'    { $cur.Text = $cur.Text + "`n" + $line }
                'Explanation' { $cur.Explanation = $cur.Explanation + "`n" + $line }
                'Option'      {
                    $idx = $cur.Options.Count - 1
                    if ($idx -ge 0) { $cur.Options[$idx] = $cur.Options[$idx] + "`n" + $line }
                    else { $cur.Text = $cur.Text + "`n" + $line }
                }
                'Answer'      {
                    $idx = $cur.Answers.Count - 1
                    if ($idx -ge 0) { $cur.Answers[$idx] = $cur.Answers[$idx] + "`n" + $line }
                    else { $cur.Text = $cur.Text + "`n" + $line }
                }
            }
            continue
        }

        switch ($kind) {
            'Question' {
                if ($cur) { $questions.Add($cur) }
                $cur = [pscustomobject]@{
                    Number           = $questions.Count + 1
                    LineNo           = $lineNo
                    Text             = $body
                    Options          = New-Object System.Collections.Generic.List[string]
                    OptionCorrect    = New-Object System.Collections.Generic.List[bool]
                    Answers          = New-Object System.Collections.Generic.List[string]
                    DeclaredType     = $null
                    Explanation      = $null
                    ExplanationCount = 0
                    Image            = $null
                    ImageCount       = 0
                    LastField        = 'Question'
                    Continuations    = 0
                }
            }
            'Option' {
                if ($cur) {
                    # The trailing asterisk is read HERE, before any continuation
                    # line is appended. Otherwise a config line ending in "*"
                    # could silently mark an option as correct.
                    $isCorrect = $body.EndsWith('*', [StringComparison]::Ordinal)
                    $text      = if ($isCorrect) { $body.Substring(0, $body.Length - 1).TrimEnd() } else { $body }
                    $cur.Options.Add($text)
                    $cur.OptionCorrect.Add($isCorrect)
                    $cur.LastField = 'Option'
                }
                else { $unknown++ }
            }
            'Answer' {
                if ($cur) { $cur.Answers.Add($body); $cur.LastField = 'Answer' }
                else { $unknown++ }
            }
            'Type' {
                if ($cur) {
                    if (-not $cur.DeclaredType) { $cur.DeclaredType = $body }
                    # Deliberately does NOT change LastField: a config block after
                    # a Type: line still belongs to the question or explanation.
                }
                else { $unknown++ }
            }
            'Explanation' {
                if ($cur) {
                    $cur.ExplanationCount++
                    if ($cur.ExplanationCount -eq 1) { $cur.Explanation = $body }
                    $cur.LastField = 'Explanation'
                }
                else { $unknown++ }
            }
            'Image' {
                if ($cur) {
                    $cur.ImageCount++
                    if ($cur.ImageCount -eq 1) { $cur.Image = $body }
                    # Deliberately does NOT change LastField - see above.
                }
                else { $unknown++ }
            }
        }
    }
    if ($cur) { $questions.Add($cur) }

    if ($unknown -gt 0) {
        $issues.Add([pscustomobject]@{
            Question = '-'
            Text     = "$unknown line(s) ignored (no recognised Question:/Option:/Answer:/Type:/Explanation:/Image: prefix)"
        })
    }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($q in $questions) {
        # Options arrive already cleaned: the trailing asterisk was read at parse
        # time, before any continuation lines were appended, so a config block
        # line ending in "*" cannot mark an option correct by accident.
        $optTexts   = New-Object System.Collections.Generic.List[string]
        $correctIdx = New-Object System.Collections.Generic.List[int]
        for ($i = 0; $i -lt $q.Options.Count; $i++) {
            $optTexts.Add($q.Options[$i])
            if ($i -lt $q.OptionCorrect.Count -and $q.OptionCorrect[$i]) { $correctIdx.Add($i + 1) }
        }

        # ---- Fill-in-the-Blank ----
        # Set with a "Type: Fill-in-the-Blank" line, and the accepted typed answers
        # are given with "Answer:" lines (or, if there are none, the Option lines).
        # Wayground expects: Question Type = "Fill-in-the-Blank", the accepted
        # answers sitting in the Option columns, and Correct Answer left BLANK.
        $declared        = if ($q.DeclaredType) { $q.DeclaredType.Trim() } else { '' }
        $isFillInTheBlank = $declared -match $FibPattern

        if ($isFillInTheBlank) {
            # @(...) around BOTH branches: a single-element List[object] unwraps to
            # a bare string when it leaves the if-expression, and .Count then fails.
            $answers = @(if ($q.Answers.Count -gt 0) { $q.Answers } else { $optTexts })
            $answerCount = $answers.Count

            $row = [ordered]@{}
            foreach ($c in $Columns) { $row[$c] = $null }
            $row['Question Text']   = $q.Text
            $row['Question Type']   = 'Fill-in-the-Blank'
            for ($k = 1; $k -le $MaxOptions; $k++) {
                $row["Option $k"] = if ($answerCount -ge $k) { $answers[$k - 1] } else { $null }
            }
            # Leave blank on purpose - that is what Wayground expects for FTB.
            $row['Correct Answer']     = $null
            $row['Time in seconds']    = $TimeSeconds
            $row['Image Link']         = $q.Image
            $row['Answer explanation'] = $q.Explanation
            $rows.Add([pscustomobject]$row)

            if ($answerCount -eq 0) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "Fill-in-the-Blank with NO accepted answers - $($q.Text)" })
            }
            if ($answerCount -gt $MaxOptions) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "Fill-in-the-Blank has $($answerCount) answers; kept the first $MaxOptions - $($q.Text)" })
            }
            if ($q.Answers.Count -gt 0 -and $q.Options.Count -gt 0) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "Fill-in-the-Blank has BOTH Answer: and Option: lines; the Answer: lines were used - $($q.Text)" })
            }
            if ($q.ImageCount -gt 1) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "$($q.ImageCount) Image: lines; kept the first - $($q.Text)" })
            }
            if ($q.ExplanationCount -eq 0) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "missing explanation - $($q.Text)" })
            }
            continue
        }

        if ($q.Answers.Count -gt 0) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "has Answer: lines but no 'Type: Fill-in-the-Blank'; they were ignored - $($q.Text)" })
        }

        $allOpts = @($optTexts)
        $kept    = @($correctIdx | Where-Object { $_ -le $MaxOptions })
        $qType   = if ($kept.Count -gt 1) { 'Checkbox' } else { 'Multiple Choice' }

        $row = [ordered]@{}
        foreach ($c in $Columns) { $row[$c] = $null }
        $row['Question Text']   = $q.Text
        $row['Question Type']   = $qType
        for ($k = 1; $k -le $MaxOptions; $k++) {
            $row["Option $k"] = if ($allOpts.Count -ge $k) { $allOpts[$k - 1] } else { $null }
        }
        $row['Correct Answer']    = ($kept -join ',')
        $row['Time in seconds']   = $TimeSeconds
        $row['Image Link']        = $q.Image
        $row['Answer explanation'] = $q.Explanation

        $rows.Add([pscustomobject]$row)

        # A declared Type: that disagrees with the asterisks. The asterisks win,
        # but the disagreement is worth knowing about.
        if ($declared -match $McPattern -and $kept.Count -gt 1) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "declared 'Multiple Choice' but $($kept.Count) answers are starred (kept as Checkbox) - $($q.Text)" })
        }
        if ($declared -match $CbPattern -and $kept.Count -le 1) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "declared 'Checkbox' but only $($kept.Count) answer(s) starred - $($q.Text)" })
        }

        # ---- QA ----
        if ($kept.Count -eq 0) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "NO CORRECT ANSWER marked - $($q.Text)" })
        }
        if ($allOpts.Count -gt $MaxOptions) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "had $($allOpts.Count) options; kept the first $MaxOptions - $($q.Text)" })
        }
        if ($allOpts.Count -lt 2) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "only $($allOpts.Count) option(s) - $($q.Text)" })
        }
        if ($q.ImageCount -gt 1) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "$($q.ImageCount) Image: lines; kept the first - $($q.Text)" })
        }
        if ($q.ExplanationCount -eq 0) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "missing explanation - $($q.Text)" })
        }

        # "Choose two" / "Choose three" in the stem vs how many options are starred.
        # A mismatch means the answer key and the question disagree.
        if ($q.Text -match '(?i)choose\s+(two|three|four|all)') {
            $want = switch ($Matches[1].ToLower()) { 'two' { 2 } 'three' { 3 } 'four' { 4 } default { -1 } }
            if ($want -gt 0 -and $kept.Count -ne $want) {
                $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "says 'Choose $($Matches[1])' but $($kept.Count) answer(s) are starred - $($q.Text)" })
            }
        }

        # Every option correct is almost always a mistake.
        if ($allOpts.Count -gt 1 -and $kept.Count -eq $allOpts.Count) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "ALL $($allOpts.Count) options are marked correct - $($q.Text)" })
        }

        # Two identical options in one question.
        $dupOpt = @($allOpts | Group-Object | Where-Object { $_.Count -gt 1 })
        if ($dupOpt.Count -gt 0) {
            $names = ($dupOpt | ForEach-Object { $_.Name }) -join '; '
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "repeated option text ($names) - $($q.Text)" })
        }

        # Long question text: usually a pasted config block. Quizizz shows these
        # in a small box, so it is worth a look rather than an error.
        if ($q.Text.Length -gt 500) {
            $issues.Add([pscustomobject]@{ Question = $q.Number; Text = "question text is $($q.Text.Length) characters (long) - $($q.Text.Substring(0,60))..." })
        }
    }

    # Repeated question text.
    # Repeats are normally INTENTIONAL - the same question placed later in the
    # quiz as a recall/spaced-repetition item. So they are counted but NOT
    # reported as problems unless -FlagDuplicates is used.
    # Each repeat is kept as its own separate question row (never merged).
    $repeats = 0
    $dupes = $questions | Group-Object { $_.Text } | Where-Object { $_.Count -gt 1 }
    foreach ($d in $dupes) {
        $repeats += $d.Count - 1
        if ($FlagDuplicates) {
            $nums = ($d.Group | ForEach-Object { $_.Number }) -join ', '
            $issues.Add([pscustomobject]@{ Question = $nums; Text = "question text appears $($d.Count) times - $($d.Name)" })
        }
    }

    [pscustomobject]@{
        Rows          = $rows
        Issues        = $issues
        WithImage     = @($rows | Where-Object { $_.'Image Link' }).Count
        Repeats       = $repeats
        Continuations = ($questions | Measure-Object -Property Continuations -Sum).Sum
    }
}

# ---------------------------------------------------------------------------
#  Minimal .xlsx writer (OOXML, no Excel install required)
# ---------------------------------------------------------------------------

function Get-ColumnLetter {
    param([int]$Number)
    $s = ''
    while ($Number -gt 0) {
        $m = ($Number - 1) % 26
        $s = [string][char](65 + $m) + $s
        $Number = [int](($Number - $m) / 26)
    }
    return $s
}

function ConvertTo-XmlText {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $v = $Value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    $v = $v.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
    return $v
}

function Write-QuizizzXlsx {
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$Path
    )

    Add-Type -AssemblyName System.IO.Compression | Out-Null

    $colCount = $Columns.Count
    $lastCol  = Get-ColumnLetter $colCount

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')
    [void]$sb.Append('<cols>')
    $widths = @(52, 16, 34, 34, 34, 34, 34, 16, 16, 42, 60)
    for ($c = 1; $c -le $colCount; $c++) {
        $w = if ($c -le $widths.Count) { $widths[$c - 1] } else { 20 }
        [void]$sb.Append("<col min=`"$c`" max=`"$c`" width=`"$w`" customWidth=`"1`"/>")
    }
    [void]$sb.Append('</cols><sheetData>')

    # Header row (style 1 = bold)
    [void]$sb.Append('<row r="1">')
    for ($c = 1; $c -le $colCount; $c++) {
        $ref = (Get-ColumnLetter $c) + '1'
        $txt = ConvertTo-XmlText $Columns[$c - 1]
        [void]$sb.Append("<c r=`"$ref`" s=`"1`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$txt</t></is></c>")
    }
    [void]$sb.Append('</row>')

    $r = 1
    foreach ($row in $Rows) {
        $r++
        [void]$sb.Append("<row r=`"$r`">")
        for ($c = 1; $c -le $colCount; $c++) {
            $name = $Columns[$c - 1]
            $val  = $row.$name
            if ($null -eq $val) { continue }
            if ($val -is [string] -and $val -eq '') { continue }

            $ref = (Get-ColumnLetter $c) + $r
            if ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal]) {
                [void]$sb.Append("<c r=`"$ref`"><v>$val</v></c>")
            }
            else {
                $txt = ConvertTo-XmlText ([string]$val)
                [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$txt</t></is></c>")
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData></worksheet>')

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
'@

    $rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@

    $workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Create a Quiz" sheetId="1" r:id="rId1"/></sheets></workbook>
'@

    $workbookRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@

    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
'@

    $parts = [ordered]@{
        '[Content_Types].xml'        = $contentTypes
        '_rels/.rels'                = $rootRels
        'xl/workbook.xml'            = $workbook
        'xl/_rels/workbook.xml.rels' = $workbookRels
        'xl/styles.xml'              = $styles
        'xl/worksheets/sheet1.xml'   = $sb.ToString()
    }

    $full = [System.IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Force }

    $stream = [System.IO.File]::Create($full)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($kv in $parts.GetEnumerator()) {
                $entry  = $zip.CreateEntry($kv.Key, [System.IO.Compression.CompressionLevel]::Optimal)
                $es     = $entry.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($kv.Value)
                    $es.Write($bytes, 0, $bytes.Length)
                }
                finally { $es.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $stream.Dispose() }
}

# ---------------------------------------------------------------------------
#  Main
# ---------------------------------------------------------------------------

$inputPaths = @(switch ($PSCmdlet.ParameterSetName) {
    'Files' { @($InputFiles | ForEach-Object { (Resolve-Path -LiteralPath $_).Path }) }
    'File'  { @((Resolve-Path -LiteralPath $InputPath).Path) }
    default { @((Resolve-Path -LiteralPath $InputFolder).Path) }
})

$isSingleFile = ($PSCmdlet.ParameterSetName -eq 'File')
$baseFolder   = if ($isSingleFile) { Split-Path -Parent $inputPaths[0] }
                elseif ($PSCmdlet.ParameterSetName -eq 'Files') { Split-Path -Parent $inputPaths[0] }
                else { $inputPaths[0] }
$resolvedInput = if ($PSCmdlet.ParameterSetName -eq 'Files') { $baseFolder } else { $inputPaths[0] }

if (-not $OutputFolder) { $OutputFolder = Join-Path $baseFolder '_QuizizzOutput' }
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# Collect input files
$files = @()
if ($PSCmdlet.ParameterSetName -eq 'Files') {
    $files = @($inputPaths | ForEach-Object { Get-Item -LiteralPath $_ } |
               Where-Object {
                   $_.Name -notlike '~$*' -and
                   $_.Name -notin @('README.txt', 'Prompt.txt', 'GenerateQuizPrompt.txt') -and
                   $_.FullName -notmatch '\\_QuizizzOutput\\'
               })
}
elseif ($isSingleFile) {
    $files = @(Get-Item -LiteralPath $resolvedInput)
}
else {
    $files = @(Get-ChildItem -LiteralPath $resolvedInput -Filter '*.txt' -File -Recurse:$Recurse |
               Where-Object {
                   $_.Name -notlike '~$*' -and
                   $_.Name -notin @('README.txt', 'Prompt.txt', 'GenerateQuizPrompt.txt') -and
                   $_.FullName -notmatch '\\_QuizizzOutput\\'
               })
}

if ($files.Count -eq 0) {
    Write-Warning "No .txt files found in $resolvedInput"
    return
}

function Get-OutputBaseName {
    param([string]$File, [string]$Root, [string]$Prefix, [bool]$Flatten)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($File)
    $base = $base -replace '([_\-\s]*raw)$', ''
    $base = $base -replace '([_\-\s]*txt)$', ''
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetFileNameWithoutExtension($File) }

    if ($Flatten) {
        $dir = Split-Path -Parent $File
        $rel = ''
        if ($dir.Length -gt $Root.Length) { $rel = $dir.Substring($Root.Length).Trim('\', '/') }
        if ($rel) { $base = ($rel -replace '[\\/]', '_') + '_' + $base }
    }
    return "$Prefix$base"
}

$report   = New-Object System.Collections.Generic.List[string]
$usedNames = @{}
$totalQuestions = 0
$totalImages    = 0
$filesWithIssues = 0
$failed = 0

$report.Add('=== Quizizz conversion report ===')
$report.Add("Input      : $resolvedInput")
$report.Add("Output     : $OutputFolder")
$report.Add("Files found: $($files.Count)")
$report.Add('')

foreach ($file in ($files | Sort-Object Name)) {
    try {
        $parsed = ConvertFrom-RawQuizText -Path $file.FullName -MaxOptions $MaxOptions -TimeSeconds $TimeSeconds -Columns $QuizizzColumns -FlagDuplicates:$FlagDuplicates -StrictLines:$StrictLines
    }
    catch {
        $failed++
        $report.Add("[FAIL] $($file.Name) - $($_.Exception.Message)")
        continue
    }

    $baseName = Get-OutputBaseName -File $file.FullName -Root $baseFolder -Prefix $NamePrefix -Flatten:$Recurse
    $name = $baseName
    $n = 2
    while ($usedNames.ContainsKey($name.ToLowerInvariant())) { $name = "$baseName($n)"; $n++ }
    $usedNames[$name.ToLowerInvariant()] = $true

    $outFile = Join-Path $OutputFolder "$name.xlsx"
    Write-QuizizzXlsx -Rows $parsed.Rows -Columns $QuizizzColumns -Path $outFile

    $qCount = $parsed.Rows.Count
    $iCount = $parsed.WithImage
    $totalQuestions += $qCount
    $totalImages    += $iCount

    # .ToArray() rather than @(...): wrapping a generic List[object] of
    # PSCustomObjects in the array operator throws on some PowerShell hosts.
    $issueList = $parsed.Issues.ToArray()
    if ($issueList.Count -gt 0) { $filesWithIssues++ }

    $marker = if ($issueList.Count -gt 0) { '[WARN]' } else { '[ OK ]' }

    # Repeats are intentional, so they are reported as a plain count, not a warning.
    $extras = @()
    if ($parsed.Repeats -gt 0) { $extras += "$($parsed.Repeats) recall repeat$(if ($parsed.Repeats -ne 1) { 's' })" }
    if ($iCount -gt 0)         { $extras += "$iCount with image" }
    if ($parsed.Continuations -gt 0) { $extras += "$($parsed.Continuations) line(s) joined" }
    $extraNote = if ($extras.Count -gt 0) { ', ' + ($extras -join ', ') } else { '' }

    $report.Add("$marker $($file.Name)  ->  $name.xlsx   ($qCount questions$extraNote)")
    foreach ($iss in $issueList) {
        $text = $iss.Text
        if ($text.Length -gt 110) { $text = $text.Substring(0, 110) + '...' }
        $report.Add("         - Q$($iss.Question): $text")
    }
}

$report.Add('')
$report.Add("Total: $($files.Count) file(s) -> $($usedNames.Count) workbook(s), $totalQuestions question(s), $totalImages with an image link.")
if ($filesWithIssues -gt 0) { $report.Add("$filesWithIssues file(s) have notes above - check them before importing to Quizizz.") }
if ($failed -gt 0) { $report.Add("$failed file(s) failed to convert.") }

$reportText = ($report -join [Environment]::NewLine)
Write-Host $reportText

$reportFile = Join-Path $OutputFolder '_conversion-report.txt'
[System.IO.File]::WriteAllText($reportFile, $reportText, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ''
Write-Host "Report saved to: $reportFile"
