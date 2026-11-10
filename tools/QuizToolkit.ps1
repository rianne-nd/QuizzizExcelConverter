#Requires -Version 5.1
<#
.SYNOPSIS
    One menu for the whole Quizizz workflow: sort by subject, convert, split,
    combine, and publish to Wayground.

.DESCRIPTION
    This is the consolidated program. It does these jobs:

      1. Convert .txt files into Quizizz-ready .xlsx workbooks
         (the same engine as "Convert My Quizzes.cmd").
      2. Sort & convert new quizzes. Drop a file named
         "Subject_WhatQuiz.txt" directly into the Quizzes folder; the toolkit
         reads the word before the first "_", moves the file into
         Quizzes\<Subject>\, converts it, opens the output folder, and copies
         that output folder path to the clipboard.
      3. Split one big quiz .txt into smaller parts of N questions.
      4. Combine split _ptN.txt files back into one quiz.
      5. Publish the new workbooks to Wayground with the optional wqu Python
         tool (menu option 6 can also upload an existing folder of .xlsx
         files).

    Double-click "Quiz Toolkit.cmd" to use the menu. You can also drag a
    file or folder onto that .cmd to start with it already selected.

    When you convert, the newest quiz file in the Quizzes folder is suggested
    first. Press Enter to use it, type a number to pick another, or press B to
    browse with File Explorer.

    The conversion engine lives in tools\Convert-Quizzes.ps1, so the Excel
    files are identical to the ones the old launcher produced. Splitting and
    combining are built in, so Python is not needed for them. The optional
    Wayground upload uses the separate "wqu" Python tool, if it is installed.

    Files that are not in the "Question:/Option:" format are moved to the
    Needs-Fixing folder instead of being converted. Files with no usable
    subject prefix are moved to Quizzes\_Unsorted.

.PARAMETER Path
    Optional file or folder to work on. Used by the menu as the current
    selection, and required by -Action Convert, Split, Combine, or Sort
    (when Sort is aimed at a specific file or folder).

.PARAMETER Action
    Menu (default), Convert, Split, Combine, Sort, or Upload. The non-menu
    actions run without prompts, which is handy for shortcuts and automation.
    Sort with no -Path scans the general Quizzes folder for files to sort.
    Upload publishes the folder (or .xlsx file) in -Path to Wayground using
    the "wqu" tool and defaults to a public assignment.

.PARAMETER QuestionsPerFile
    Split size. Default 10 questions per part.

.PARAMETER TailThreshold
    If the leftover is this many questions or fewer, keep them together in
    one final file instead of making a tiny one. Default 15.

.PARAMETER Recurse
    For -Action Convert on a folder: also convert .txt files in sub-folders.

.PARAMETER NoOpen
    Automation switch for auto-sort (the Sort action, menu option 2, and the
    general-folder branch of menu option 1): do not open the output folder.
    The clipboard copy still happens unless -NoClipboard is also used.

.PARAMETER NoClipboard
    Automation switch for auto-sort: do not copy the output folder path to
    the clipboard.

.EXAMPLE
    .\Quiz Toolkit.cmd
    Open the menu (the normal way to use this).

.EXAMPLE
    .\Quiz Toolkit.cmd -Action Sort
    Sort every new .txt sitting directly in the Quizzes folder, convert each
    one, open the output folders, and copy the last output path to the
    clipboard.

.EXAMPLE
    .\Quiz Toolkit.cmd -Action Convert -Path "C:\...\DataCom"
    Convert a folder without opening the menu.

.EXAMPLE
    .\Quiz Toolkit.cmd -Action Split -Path "C:\...\Quiz.txt" -QuestionsPerFile 20
    Split a quiz into parts of 20 questions.

.EXAMPLE
    .\Quiz Toolkit.cmd -Action Combine -Path "C:\...\Quiz_pt1.txt"
    Combine one quiz family back into a single file.

.EXAMPLE
    .\Quiz Toolkit.cmd -Action Upload -Path "C:\...\Quizzes\DataComs\_QuizizzOutput"
    Publish every .xlsx in that folder to Wayground (public assignment).
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [ValidateSet('Menu', 'Convert', 'Split', 'Combine', 'Sort', 'Upload')]
    [string]$Action = 'Menu',

    [int]$QuestionsPerFile = 10,

    [int]$TailThreshold = 15,

    [switch]$Recurse,

    [switch]$NoOpen,

    [switch]$NoClipboard
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot        = Split-Path -Parent $MyInvocation.MyCommand.Path
$WorkspaceRoot     = Split-Path -Parent $ScriptRoot
$QuizzesRoot       = Join-Path $WorkspaceRoot 'Quizzes'
$Converter         = Join-Path $ScriptRoot 'Convert-Quizzes.ps1'
$Utf8NoBom         = New-Object System.Text.UTF8Encoding($false)
$WquMissingMessage = "Python 'wqu' tool not found. Run pip install -e . in the wayground-quiz-uploader folder to enable auto-uploading."

# ---------------------------------------------------------------------------
#  Small console helpers
# ---------------------------------------------------------------------------

function Clear-Screen {
    try { Clear-Host } catch { }
}

function Write-Banner {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '   QUIZ TOOLKIT' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
}

function Read-MenuChoice {
    param(
        [string]$Prompt,
        [string[]]$Valid
    )
    while ($true) {
        $answer = Read-Host $Prompt
        if ($null -eq $answer) { return '0' }
        $answer = $answer.Trim()
        if ($Valid -contains $answer) { return $answer }
        Write-Host ("  Please type one of: " + ($Valid -join ', ')) -ForegroundColor Yellow
    }
}

function ConvertTo-CleanPath {
    param([string]$Value)
    if ($null -eq $Value) { return '' }

    $p = $Value.Trim()

    # Explorer and drag-and-drop often wrap the path in quotes. Remove any
    # leading/trailing quote characters, even if only one side survived.
    $p = $p.Trim([char[]]@('"', "'")).Trim()

    $p = [Environment]::ExpandEnvironmentVariables($p)

    if ($p -eq '~') { return $HOME }
    if ($p.StartsWith('~\') -or $p.StartsWith('~/')) {
        return (Join-Path $HOME $p.Substring(2))
    }

    # A trailing slash is harmless, but a path ending in \" can confuse the
    # command line. Drop it, except when the path is a drive or share root.
    if ($p -eq '\' -or $p -eq '/') { return $p }
    if ($p -match '^[A-Za-z]:[\\/]$') { return $p }
    if ($p -match '^\\\\[^\\]+\\[^\\]+[\\/]$') { return $p }
    $p = $p.TrimEnd([char[]]@('\', '/'))

    return $p
}

function Resolve-ExistingPath {
    param(
        [string]$Value,
        [string]$Label = 'path'
    )
    $clean = ConvertTo-CleanPath $Value
    if ([string]::IsNullOrWhiteSpace($clean)) {
        throw "No $Label was provided."
    }
    if (-not (Test-Path -LiteralPath $clean)) {
        throw "Not found: $clean"
    }
    return (Resolve-Path -LiteralPath $clean).Path
}

function Show-FilePicker {
    param(
        [string]$Title = 'Select a quiz .txt file',
        [string]$InitialDirectory
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = $Title
        $dialog.Filter = 'Quiz text files (*.txt)|*.txt|All files (*.*)|*.*'
        $dialog.Multiselect = $false
        $dialog.CheckFileExists = $true
        $dialog.RestoreDirectory = $true

        if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
            $dialog.InitialDirectory = $InitialDirectory
        }
        elseif (Test-Path -LiteralPath $QuizzesRoot) {
            $dialog.InitialDirectory = $QuizzesRoot
        }
        elseif (Test-Path -LiteralPath $WorkspaceRoot) {
            $dialog.InitialDirectory = $WorkspaceRoot
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.FileName
        }
        return $null
    }
    catch {
        Write-Host "  Could not open the file picker: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Show-FolderPicker {
    param(
        [string]$InitialDirectory,
        [string]$Description = 'Select the folder that holds the quiz .txt files'
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        $dialog.ShowNewFolderButton = $false

        if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
            $dialog.SelectedPath = $InitialDirectory
        }
        elseif (Test-Path -LiteralPath $QuizzesRoot) {
            $dialog.SelectedPath = $QuizzesRoot
        }
        elseif (Test-Path -LiteralPath $WorkspaceRoot) {
            $dialog.SelectedPath = $WorkspaceRoot
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        return $null
    }
    catch {
        Write-Host "  Could not open the folder picker: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Read-Path {
    param(
        [string]$Prompt,
        [switch]$AllowEmpty
    )
    while ($true) {
        Write-Host ''
        $answer = Read-Host $Prompt
        if ($null -eq $answer) {
            if ($AllowEmpty) { return $null }
            throw 'No input was provided.'
        }
        $answer = $answer.Trim()
        if ($answer -eq '') {
            if ($AllowEmpty) { return $null }
            Write-Host '  Nothing was entered. Please try again.' -ForegroundColor Yellow
            continue
        }

        $lower = $answer.ToLowerInvariant()
        if ($lower -eq 'b' -or $lower -eq 'browse') {
            $picked = Show-FilePicker
            if ($picked) { return $picked }
            continue
        }
        if ($lower -eq 'f' -or $lower -eq 'folder') {
            $picked = Show-FolderPicker
            if ($picked) { return $picked }
            continue
        }

        try {
            return (Resolve-ExistingPath $answer)
        }
        catch {
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Get-RecentQuizFiles {
    param([int]$Count = 9)

    $searchRoots = @()
    if (Test-Path -LiteralPath $QuizzesRoot) {
        $searchRoots += $QuizzesRoot
    }
    else {
        $searchRoots += $WorkspaceRoot
    }

    $files = New-Object System.Collections.Generic.List[object]
    foreach ($searchRoot in $searchRoots) {
        $found = @(Get-ChildItem -LiteralPath $searchRoot -Recurse -Filter '*.txt' -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.FullName -notmatch '\\_QuizizzOutput\\' -and
                $_.FullName -notmatch '\\_backups\\' -and
                $_.FullName -notmatch '\\tests\\' -and
                $_.FullName -notmatch '\\tools\\' -and
                $_.FullName -notmatch '\\powerquery\\' -and
                $_.FullName -notmatch '\\Needs-Fixing\\' -and
                $_.FullName -notmatch '\\_Unsorted\\' -and
                $_.BaseName -notmatch '_pt\d+$' -and
                $_.Name -notin @('README.txt', 'Prompt.txt', 'GenerateQuizPrompt.txt') -and
                $_.Name -notlike '~$*'
            })
        foreach ($f in $found) { $files.Add($f) }
    }

    return @($files | Sort-Object LastWriteTime -Descending | Select-Object -First $Count)
}

function Show-RecentQuizzes {
    param([object[]]$Files)

    if ($null -eq $Files -or $Files.Count -eq 0) {
        Write-Host '  No quiz files found in the Quizzes folder yet.' -ForegroundColor Yellow
        return
    }

    $rows = @()
    foreach ($f in $Files) {
        $rel = $f.FullName
        if ($rel.StartsWith($WorkspaceRoot, [StringComparison]::OrdinalIgnoreCase)) {
            $rel = $rel.Substring($WorkspaceRoot.Length).TrimStart('\', '/')
        }
        $rows += [pscustomobject]@{
            Rel  = $rel
            When = $f.LastWriteTime.ToString('MMM d, h:mm tt')
        }
    }

    $width = 0
    foreach ($row in $rows) {
        if ($row.Rel.Length -gt $width) { $width = $row.Rel.Length }
    }
    if ($width -gt 55) { $width = 55 }
    if ($width -lt 20) { $width = 20 }

    Write-Host '  Most recent quiz files:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $rel = $rows[$i].Rel
        if ($rel.Length -gt $width) {
            $rel = '...' + $rel.Substring($rel.Length - $width + 3)
        }
        Write-Host ("    [{0}]  {1}  {2}" -f ($i + 1), $rel.PadRight($width), $rows[$i].When)
    }
}

function Read-QuizTarget {
    param([string]$Prompt = 'Which quiz?')

    $recent = @(Get-RecentQuizFiles -Count 9)
    $browseStart = if ($recent.Count -gt 0) { Split-Path -Parent $recent[0].FullName } else { $QuizzesRoot }

    Write-Host ''
    Show-RecentQuizzes -Files $recent
    Write-Host ''
    Write-Host '  Press Enter for the newest, type a number, B to browse files, F for a folder,' -ForegroundColor DarkGray
    Write-Host '  or drag a file/folder here.' -ForegroundColor DarkGray

    while ($true) {
        $answer = Read-Host $Prompt
        if ($null -eq $answer) { return $null }
        $answer = $answer.Trim()

        if ($answer -eq '') {
            if ($recent.Count -gt 0) { return $recent[0].FullName }
            Write-Host '  There is no recent file to choose. Type B to browse, F for a folder, or drag a file here.' -ForegroundColor Yellow
            continue
        }

        if ($answer -match '^\d+$') {
            $n = 0
            if ([int]::TryParse($answer, [ref]$n) -and $n -ge 1 -and $n -le $recent.Count) {
                return $recent[$n - 1].FullName
            }
            if ($recent.Count -gt 0) {
                Write-Host "  Type a number from 1 to $($recent.Count), or press Enter for the newest." -ForegroundColor Yellow
            }
            else {
                Write-Host '  There is no recent file with that number.' -ForegroundColor Yellow
            }
            continue
        }

        $lower = $answer.ToLowerInvariant()
        if ($lower -eq 'b' -or $lower -eq 'browse') {
            $picked = Show-FilePicker -InitialDirectory $browseStart
            if ($picked) { return $picked }
            continue
        }
        if ($lower -eq 'f' -or $lower -eq 'folder') {
            $picked = Show-FolderPicker
            if ($picked) { return $picked }
            continue
        }

        try {
            return (Resolve-ExistingPath $answer)
        }
        catch {
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Read-Int {
    param(
        [string]$Prompt,
        [int]$Default,
        [int]$Minimum = 0
    )
    while ($true) {
        $answer = Read-Host $Prompt
        if ($null -eq $answer -or $answer.Trim() -eq '') { return $Default }

        $n = 0
        if ([int]::TryParse($answer.Trim(), [ref]$n) -and $n -ge $Minimum) {
            return $n
        }
        Write-Host "  Please type a whole number ($Minimum or more), or press Enter for $Default." -ForegroundColor Yellow
    }
}

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $false
    )
    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }

    while ($true) {
        $answer = Read-Host "$Prompt $hint"
        if ($null -eq $answer -or $answer.Trim() -eq '') { return $Default }

        switch ($answer.Trim().ToLowerInvariant()) {
            'y'   { return $true }
            'yes' { return $true }
            'n'   { return $false }
            'no'  { return $false }
            default { Write-Host '  Please answer y or n.' -ForegroundColor Yellow }
        }
    }
}

function Wait-Menu {
    Write-Host ''
    Read-Host 'Press Enter to return to the menu' | Out-Null
}

function Test-WquInstalled {
    if (Get-Command wqu -ErrorAction SilentlyContinue) {
        return $true
    }

    Write-Host ''
    Write-Host "  $WquMissingMessage" -ForegroundColor Yellow
    return $false
}

function Read-AssignArguments {
    $assignToNaic = Read-YesNo -Prompt "Assign to class 'Naic'? (N for public 1-day deadline)" -Default $true
    if ($assignToNaic) {
        return @('--assign', 'Naic')
    }
    return @('--no-assign')
}

function Invoke-WquPublish {
    param(
        [string[]]$Paths,
        [string[]]$AssignArguments = @()
    )

    if ($null -eq $Paths -or $Paths.Count -eq 0) { return }
    if (-not (Test-WquInstalled)) { return }

    $absolutePaths = New-Object System.Collections.Generic.List[string]
    foreach ($pathValue in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($pathValue)) { continue }

        try {
            $resolved = (Resolve-Path -LiteralPath $pathValue -ErrorAction Stop).Path
            if (-not $absolutePaths.Contains($resolved)) {
                $absolutePaths.Add($resolved) | Out-Null
            }
        }
        catch {
            Write-Host "  Skipping upload path (not found): $pathValue" -ForegroundColor Yellow
        }
    }

    if ($absolutePaths.Count -eq 0) {
        Write-Host '  No uploadable files or folders were left.' -ForegroundColor Yellow
        return
    }

    if ($null -eq $AssignArguments) { $AssignArguments = @() }

    $quoted = @()
    foreach ($p in $absolutePaths) { $quoted += '"' + $p + '"' }

    $display = "  wqu publish $($quoted -join ' ')"
    if ($AssignArguments.Count -gt 0) { $display += " $($AssignArguments -join ' ')" }

    Write-Host ''
    Write-Host '  Starting Wayground upload...' -ForegroundColor Cyan
    Write-Host $display -ForegroundColor DarkGray

    # Run synchronously (not Start-Process) so wqu's live tree stays in this
    # window and the user can answer any follow-up prompt from the Python tool.
    $exitCode = 0
    & wqu publish @absolutePaths @AssignArguments

    # $LASTEXITCODE is only created after a native command. Check for it so
    # strict mode stays happy if wqu is a PowerShell shim instead of an .exe.
    $lastExit = Get-Variable -Name LASTEXITCODE -ErrorAction SilentlyContinue
    if ($null -ne $lastExit -and $null -ne $lastExit.Value) {
        $exitCode = [int]$lastExit.Value
    }

    Write-Host ''
    if ($exitCode -eq 0) {
        Write-Host '  Wayground upload finished.' -ForegroundColor Green
    }
    else {
        Write-Host "  wqu stopped with exit code $exitCode. Read the messages above." -ForegroundColor Yellow
        Write-Host "  If the session expired, run 'wqu login' and try again." -ForegroundColor DarkGray
    }
}

function Read-UploadFolder {
    param([string]$DefaultFolder)

    while ($true) {
        Write-Host ''
        if ($DefaultFolder -and (Test-Path -LiteralPath $DefaultFolder)) {
            Write-Host "  Default: $DefaultFolder" -ForegroundColor DarkGray
        }
        Write-Host '  Press Enter to use the default, type B to browse, or drag a folder here.' -ForegroundColor DarkGray

        $answer = Read-Host 'Folder with the .xlsx files to upload'
        if ($null -eq $answer) { return $null }
        $answer = $answer.Trim()

        if ($answer -eq '') {
            if ($DefaultFolder -and (Test-Path -LiteralPath $DefaultFolder)) {
                return (Resolve-Path -LiteralPath $DefaultFolder).Path
            }
            Write-Host "  The default folder does not exist yet: $DefaultFolder" -ForegroundColor Yellow
            continue
        }

        $lower = $answer.ToLowerInvariant()
        if ($lower -eq 'b' -or $lower -eq 'browse') {
            $picked = Show-FolderPicker -InitialDirectory $DefaultFolder -Description 'Select a folder that contains .xlsx quiz files'
            if ($picked) { return $picked }
            continue
        }

        try {
            return (Resolve-ExistingPath $answer)
        }
        catch {
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ---------------------------------------------------------------------------
#  File helpers
# ---------------------------------------------------------------------------

function Read-Utf8Text {
    param([string]$FilePath)
    return [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text {
    param(
        [string]$FilePath,
        [string]$Text
    )
    # Write Windows line endings, like the old Python tools did, so the
    # files open correctly in Notepad and every other Windows editor.
    $normalized = $Text.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllText($FilePath, $normalized, $script:Utf8NoBom)
}

# ---------------------------------------------------------------------------
#  Splitter (same rules as the old Splitter.py)
# ---------------------------------------------------------------------------

function Split-QuizText {
    param([string]$Text)

    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $questions  = New-Object System.Collections.Generic.List[string]
    $preamble   = New-Object System.Collections.Generic.List[string]
    $current    = $null

    foreach ($line in $normalized.Split([char]10)) {
        if ($line.Trim() -match '^Question:') {
            if ($null -ne $current) {
                $questions.Add(($current -join "`n").Trim())
            }
            $current = New-Object System.Collections.Generic.List[string]
            $current.Add($line)
        }
        elseif ($null -eq $current) {
            if ($line.Trim() -ne '') { $preamble.Add($line) }
        }
        else {
            $current.Add($line)
        }
    }

    if ($null -ne $current) {
        $questions.Add(($current -join "`n").Trim())
    }

    return [pscustomobject]@{
        Questions = @($questions | Where-Object { $_ -ne '' })
        Preamble  = $preamble
    }
}

function Split-IntoBlocks {
    param([string]$Text)

    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
    if ($normalized -eq '') { return @() }

    $parts = [regex]::Split($normalized, "\n\s*\n+")
    return @($parts | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
}

function Get-Chunks {
    param(
        [object[]]$Items,
        [int]$Size,
        [int]$TailThreshold
    )

    $chunks = New-Object System.Collections.Generic.List[object]
    $i = 0
    $n = $Items.Count

    while ($i -lt $n) {
        $remaining = $n - $i

        if ($remaining -le $TailThreshold) {
            $chunks.Add(@($Items[$i..($n - 1)]))
            break
        }

        $last = $i + $Size - 1
        if ($last -ge $n) { $last = $n - 1 }

        $chunks.Add(@($Items[$i..$last]))
        $i += $Size
    }

    return $chunks
}

function Invoke-SplitFile {
    param(
        [System.IO.FileInfo]$File,
        [int]$Size,
        [int]$TailThreshold
    )

    $result = [pscustomobject]@{
        Success = $false
        Folder  = $File.DirectoryName
        Count   = 0
    }

    try {
        $text = Read-Utf8Text -FilePath $File.FullName
    }
    catch {
        Write-Host "  Error reading $($File.Name): $($_.Exception.Message)" -ForegroundColor Red
        return $result
    }

    $parsed    = Split-QuizText -Text $text
    $questions = @($parsed.Questions)

    if ($questions.Count -eq 0) {
        $questions = @(Split-IntoBlocks -Text $text)
        if ($questions.Count -eq 0) {
            Write-Host "  Skipping $($File.Name): no valid question blocks found." -ForegroundColor Yellow
            return $result
        }
        Write-Host "  NOTE: $($File.Name) has no 'Question:' lines - falling back to blank-line splitting ($($questions.Count) blocks)." -ForegroundColor Yellow
    }

    if ($parsed.Preamble.Count -gt 0) {
        $sample = $parsed.Preamble[0]
        if ($sample.Length -gt 70) { $sample = $sample.Substring(0, 70) }
        Write-Host "  NOTE: $($parsed.Preamble.Count) line(s) before the first 'Question:' line were ignored, e.g.: $sample" -ForegroundColor Yellow
    }

    $chunks = @(Get-Chunks -Items $questions -Size $Size -TailThreshold $TailThreshold)

    if ($chunks.Count -eq 1 -and $questions.Count -le $Size) {
        Write-Host "  $($File.Name): only $($questions.Count) question(s); nothing to split." -ForegroundColor DarkGray
        return $result
    }

    Write-Host ''
    Write-Host "  File: $($File.Name)"
    Write-Host "  Found $($questions.Count) question(s). Creating $($chunks.Count) file(s):"

    $start = 1
    for ($i = 0; $i -lt $chunks.Count; $i++) {
        $piece      = @($chunks[$i])
        $partNumber = $i + 1
        $outName    = "$($File.BaseName)_pt$partNumber$($File.Extension)"
        $outPath    = Join-Path $File.DirectoryName $outName
        $outText    = (($piece -join "`n`n") + "`n")

        Write-Utf8Text -FilePath $outPath -Text $outText

        $end = $start + $piece.Count - 1
        Write-Host "    -> Created: $outName ($($piece.Count) questions, #$start-#$end)" -ForegroundColor Green
        $start = $end + 1
    }

    $result.Success = $true
    $result.Count   = $chunks.Count
    return $result
}

function Invoke-Split {
    param(
        [string]$Target,
        [int]$Size,
        [int]$TailThreshold
    )

    $item  = Get-Item -LiteralPath $Target
    $files = @()

    if ($item.PSIsContainer) {
        $files = @(Get-ChildItem -LiteralPath $item.FullName -Filter '*.txt' -File |
                   Where-Object { $_.BaseName -notmatch '_pt\d+$' } |
                   Sort-Object Name)

        if ($files.Count -eq 0) {
            Write-Host '  No .txt files found in that folder.' -ForegroundColor Yellow
            return [pscustomobject]@{ Success = $false; Folder = $item.FullName; Folders = @(); SplitCount = 0 }
        }
        Write-Host "  Found $($files.Count) file(s) to check."
    }
    else {
        $files = @($item)
    }

    $anySplit   = $false
    $totalParts = 0
    $folders    = New-Object System.Collections.Generic.List[string]

    foreach ($file in $files) {
        $workingFile = $file

        # A quiz dropped in the general Quizzes\ folder is filed into its
        # subject folder before it is split, so the parts are created next to
        # the original (Quizzes\<Subject>\) instead of in the general folder
        # and its output stays in that subject's _QuizizzOutput.
        if (Test-IsGeneralFolderPath -FolderPath $file.DirectoryName) {
            $sort = Invoke-AutoSort -Target $file.FullName -SkipConversion -PassThruResult

            if ($sort.SortedFiles.Count -gt 0) {
                $workingFile = Get-Item -LiteralPath $sort.SortedFiles[0]
                Write-Host "    - Filed into: $([System.IO.Path]::GetFileName($workingFile.DirectoryName))\" -ForegroundColor Green
            }
            elseif (-not (Test-Path -LiteralPath $file.FullName)) {
                Write-Host '    - Moved out of the general folder while sorting; split skipped for this file.' -ForegroundColor Yellow
                continue
            }
        }

        $r = Invoke-SplitFile -File $workingFile -Size $Size -TailThreshold $TailThreshold
        if ($r.Success) {
            $anySplit = $true
            $totalParts += $r.Count
        }

        if (-not $folders.Contains($workingFile.DirectoryName)) {
            $folders.Add($workingFile.DirectoryName) | Out-Null
        }
    }

    $folder = if ($folders.Count -gt 0) { $folders[$folders.Count - 1] }
              elseif ($item.PSIsContainer) { $item.FullName }
              else { $item.DirectoryName }

    return [pscustomobject]@{
        Success    = $anySplit
        Folder     = $folder
        Folders    = @($folders)
        SplitCount = $totalParts
    }
}

# ---------------------------------------------------------------------------
#  Combiner (same rules as the old combiner.py)
# ---------------------------------------------------------------------------

function Invoke-Combine {
    param([string]$Target)

    $item      = Get-Item -LiteralPath $Target
    $searchDir = ''
    $outDir    = ''
    $onlyBase  = $null

    if ($item.PSIsContainer) {
        $searchDir = $item.FullName
        if ($item.Name.ToLowerInvariant() -eq 'split questions') {
            $outDir = Join-Path $item.Parent.FullName 'Combined Quizzes'
        }
        else {
            $outDir = Join-Path $searchDir 'Combined Quizzes'
        }
    }
    else {
        $searchDir = $item.DirectoryName
        $m = [regex]::Match($item.BaseName, '^(.*)_pt(\d+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $m.Success) {
            Write-Host "  Error: '$($item.Name)' does not match the split pattern (e.g., name_pt1.txt)." -ForegroundColor Yellow
            return [pscustomobject]@{ Success = $false; OutputFolder = $searchDir; Combined = 0 }
        }
        $onlyBase = $m.Groups[1].Value
        $outDir   = $searchDir
    }

    $files  = @(Get-ChildItem -LiteralPath $searchDir -Filter '*.txt' -File)
    $groups = @{}

    foreach ($f in $files) {
        $m = [regex]::Match($f.BaseName, '^(.*)_pt(\d+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            $base = $m.Groups[1].Value
            $part = [int]$m.Groups[2].Value

            if (-not $groups.ContainsKey($base)) {
                $groups[$base] = New-Object System.Collections.Generic.List[object]
            }
            $groups[$base].Add([pscustomobject]@{ Part = $part; File = $f })
        }
    }

    if ($null -ne $onlyBase) {
        $filtered = @{}
        foreach ($key in $groups.Keys) {
            if ($key -eq $onlyBase) { $filtered[$key] = $groups[$key] }
        }
        $groups = $filtered
    }

    if ($groups.Count -eq 0) {
        Write-Host "  No split files (_pt1, _pt2, ...) found in: $searchDir" -ForegroundColor Yellow
        return [pscustomobject]@{ Success = $false; OutputFolder = $outDir; Combined = 0 }
    }

    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    Write-Host "  Found $($groups.Count) quiz family/families to combine:"
    $combined = 0

    foreach ($base in ($groups.Keys | Sort-Object)) {
        $parts  = @($groups[$base] | Sort-Object Part)
        $blocks = New-Object System.Collections.Generic.List[string]

        foreach ($p in $parts) {
            try {
                $content = (Read-Utf8Text -FilePath $p.File.FullName).Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                if ($content -ne '') { $blocks.Add($content) }
            }
            catch {
                Write-Host "  Error reading $($p.File.Name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        if ($blocks.Count -gt 0) {
            $outName = "$($base)_combined.txt"
            $outFile = Join-Path $outDir $outName
            Write-Utf8Text -FilePath $outFile -Text (($blocks -join "`n`n") + "`n")
            Write-Host "    -> Merged $($parts.Count) part(s) into: $outName" -ForegroundColor Green
            $combined++
        }
    }

    return [pscustomobject]@{
        Success      = ($combined -gt 0)
        OutputFolder = $outDir
        Combined     = $combined
    }
}

# ---------------------------------------------------------------------------
#  Converter wrapper
# ---------------------------------------------------------------------------

function Get-ExcelFileSnapshot {
    param([string]$Folder)

    $snapshot = @{}
    if (-not (Test-Path -LiteralPath $Folder)) { return $snapshot }

    foreach ($file in @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -notlike '~$*' })) {
        $snapshot[$file.FullName.ToLowerInvariant()] = "$($file.Length)|$($file.LastWriteTimeUtc.Ticks)"
    }
    return $snapshot
}

function Get-NewExcelFiles {
    param(
        [string]$Folder,
        [hashtable]$Before
    )

    $created = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $Folder)) { return @() }

    foreach ($file in @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -notlike '~$*' })) {
        $key = $file.FullName.ToLowerInvariant()
        $signature = "$($file.Length)|$($file.LastWriteTimeUtc.Ticks)"

        if (-not $Before.ContainsKey($key) -or $Before[$key] -ne $signature) {
            $created.Add($file.FullName) | Out-Null
        }
    }

    return @($created | Sort-Object)
}

function Get-ConversionOutputFolder {
    param([string]$Target)

    $item = Get-Item -LiteralPath $Target
    if ($item.PSIsContainer) {
        return (Join-Path $item.FullName '_QuizizzOutput')
    }
    return (Join-Path $item.DirectoryName '_QuizizzOutput')
}

function Invoke-Conversion {
    param(
        [string]$Target,
        [switch]$RecurseSubfolders
    )

    if (-not (Test-Path -LiteralPath $Converter)) {
        throw "Could not find the converter at: $Converter"
    }

    $item = Get-Item -LiteralPath $Target
    $outputFolder = if ($item.PSIsContainer) {
        Join-Path $item.FullName '_QuizizzOutput'
    }
    else {
        Join-Path $item.DirectoryName '_QuizizzOutput'
    }

    Write-Host ''
    Write-Host '  Converting to Quizizz Excel files...' -ForegroundColor Cyan
    Write-Host ''

    # Snapshot the folder before the run so we can return exactly the
    # workbooks this conversion created or rewrote - not older files that
    # were already sitting in _QuizizzOutput.
    $before = Get-ExcelFileSnapshot -Folder $outputFolder

    if ($item.PSIsContainer) {
        $arguments = @{ InputFolder = $item.FullName }
        if ($RecurseSubfolders) { $arguments['Recurse'] = $true }
        & $Converter @arguments | Out-Null
    }
    else {
        & $Converter -InputPath $item.FullName | Out-Null
    }

    return @(Get-NewExcelFiles -Folder $outputFolder -Before $before)
}

function Invoke-SmartUpload {
    param([string[]]$Files)

    $created = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    foreach ($file in @($Files)) {
        if ([string]::IsNullOrWhiteSpace($file)) { continue }

        try {
            $resolved = (Resolve-Path -LiteralPath $file -ErrorAction Stop).Path
        }
        catch {
            continue
        }

        $key = $resolved.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $created.Add($resolved) | Out-Null
    }

    if ($created.Count -eq 0) { return }

    if (-not (Test-WquInstalled)) { return }

    $assignArguments = @(Read-AssignArguments)

    if ($created.Count -eq 1) {
        $name = [System.IO.Path]::GetFileName($created[0])
        if (-not (Read-YesNo -Prompt "Upload '$name' to Wayground now?" -Default $true)) {
            Write-Host '  Skipped the Wayground upload.' -ForegroundColor DarkGray
            return
        }

        Invoke-WquPublish -Paths @($created[0]) -AssignArguments $assignArguments
        return
    }

    Write-Host ''
    Write-Host '  Newly created Excel files:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $created.Count; $i++) {
        Write-Host ("    [{0}]  {1}" -f ($i + 1), $created[$i])
    }

    while ($true) {
        $answer = Read-Host 'Upload these to Wayground? [A]ll, [S]elect specific, or [N]one:'
        if ($null -eq $answer) { return }

        switch ($answer.Trim().ToLowerInvariant()) {
            'a' {
                Invoke-WquPublish -Paths @($created) -AssignArguments $assignArguments
                return
            }
            'all' {
                Invoke-WquPublish -Paths @($created) -AssignArguments $assignArguments
                return
            }
            's' {
                $selection = Read-Host 'Enter numbers separated by commas (e.g., 1,3)'
                if ($null -eq $selection) { return }

                $pickedNumbers = New-Object System.Collections.Generic.List[int]
                foreach ($part in $selection.Split(',')) {
                    $text = $part.Trim()
                    if ($text -eq '') { continue }

                    $number = 0
                    if ([int]::TryParse($text, [ref]$number) -and
                        $number -ge 1 -and $number -le $created.Count) {
                        if (-not $pickedNumbers.Contains($number)) {
                            $pickedNumbers.Add($number) | Out-Null
                        }
                    }
                    else {
                        Write-Host "  Ignoring '$text' (expected 1 to $($created.Count))." -ForegroundColor Yellow
                    }
                }

                if ($pickedNumbers.Count -eq 0) {
                    Write-Host '  No valid numbers were entered; skipping the upload.' -ForegroundColor Yellow
                    return
                }

                $pickedFiles = New-Object System.Collections.Generic.List[string]
                foreach ($number in @($pickedNumbers | Sort-Object)) {
                    $pickedFiles.Add($created[$number - 1]) | Out-Null
                }

                Invoke-WquPublish -Paths @($pickedFiles) -AssignArguments $assignArguments
                return
            }
            'n' {
                Write-Host '  Skipped the Wayground upload.' -ForegroundColor DarkGray
                return
            }
            'none' {
                Write-Host '  Skipped the Wayground upload.' -ForegroundColor DarkGray
                return
            }
            default {
                Write-Host '  Please type A, S, or N.' -ForegroundColor Yellow
            }
        }
    }
}

# ---------------------------------------------------------------------------
#  Auto-sort: "Subject_WhatQuiz.txt" -> Quizzes\<Subject>\ + conversion
# ---------------------------------------------------------------------------

function Set-ClipboardText {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }

    try {
        Set-Clipboard -Value $Value -ErrorAction Stop
        return $true
    }
    catch {
        try {
            $Value | clip.exe
            return $true
        }
        catch {
            return $false
        }
    }
}

function Get-AvailableFileName {
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Extension
    )

    $candidate = Join-Path $Folder ($BaseName + $Extension)
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $i = 2
    while ($true) {
        $candidate = Join-Path $Folder ("$BaseName($i)$Extension")
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        $i++
    }
}

function Test-IsGeneralFolderPath {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $false }
    $a = $FolderPath.TrimEnd('\', '/')
    $b = $QuizzesRoot.TrimEnd('\', '/')
    return ($a -ieq $b)
}

function Get-QuizFormatInfo {
    param([string]$FilePath)

    $info = [pscustomobject]@{
        Valid    = 0
        Malformed = 0
        Lines    = 0
        ReadError = $null
    }

    try {
        $lines = [System.IO.File]::ReadAllLines($FilePath)
    }
    catch {
        $info.ReadError = $_.Exception.Message
        return $info
    }

    $info.Lines = $lines.Count
    foreach ($line in $lines) {
        # Match the converter's own cleanup: zero-width characters and
        # non-breaking spaces must not make a valid file look malformed.
        $clean = $line -replace '[\u200B-\u200D\uFEFF]', ''
        $clean = $clean -replace '\u00A0', ' '
        if ($clean -match '^\s*Question\s*:') { $info.Valid++ }
        elseif ($clean -match '^\s*Question\s+[^:]') { $info.Malformed++ }
    }
    return $info
}

function Get-CleanSubjectName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $name = $Value.Trim()
    $name = $name -replace '[<>:"/\\|?*]', ''
    $name = $name.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    return $name
}

function Resolve-SubjectFolder {
    <#
        Reads the text before the first "_" in the file name and turns it
        into the matching subject folder under Quizzes\.
        Returns $null when there is no usable subject prefix.
    #>
    param([string]$FileName)

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $idx  = $base.IndexOf('_')
    if ($idx -le 0) { return $null }

    $prefix = Get-CleanSubjectName $base.Substring(0, $idx)
    if (-not $prefix) { return $null }

    # An existing folder wins, so "datacoms_Mod1.txt" lands in "DataComs".
    if (Test-Path -LiteralPath $QuizzesRoot) {
        $existing = Get-ChildItem -LiteralPath $QuizzesRoot -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -ieq $prefix } |
                    Select-Object -First 1
        if ($existing) { return $existing.FullName }
    }

    # Do not invent a subject folder out of a module/checkpoint-style name.
    if ($prefix -match '^(?i)(mod|module|modules|checkpoint|checkoint|quiz|quizzes|exam|exams|lt|lq|raw|test|new|intro|introduction)\d*$') {
        return $null
    }
    if ($prefix -match '^\d+$') { return $null }

    return (Join-Path $QuizzesRoot $prefix)
}

function Invoke-AutoSortCore {
    <#
        Sorts files by their "Subject_" prefix:
          1. find .txt files (the general Quizzes folder by default),
          2. read the prefix before the first "_",
          3. move the file into Quizzes\<Subject>\ (creating the folder if it
             is genuinely new),
          4. convert it to Quizizz Excel,
          5. open the output folder(s), and copy the last output path to the
             clipboard.

        Wrong-format files go to Needs-Fixing. Files without a usable prefix
        go to Quizzes\_Unsorted. Same-name collisions keep both files.

        -SkipConversion files the quiz into its subject folder without
        converting it. The menu uses this before splitting a quiz that was
        dropped in the general Quizzes\ folder.
    #>
    param(
        [string]$Target,
        [switch]$SkipConversion
    )

    $result = [pscustomobject]@{
        Success       = $false
        Sorted        = 0
        Skipped       = 0
        NeedsFixing   = 0
        Failed        = 0
        SortedFiles   = New-Object System.Collections.Generic.List[string]
        OutputFolders = New-Object System.Collections.Generic.List[string]
        CreatedFiles  = New-Object System.Collections.Generic.List[string]
        CopiedPath    = $null
    }

    $files = @()

    if ([string]::IsNullOrWhiteSpace($Target)) {
        if (-not (Test-Path -LiteralPath $QuizzesRoot)) {
            Write-Host "  The Quizzes folder was not found: $QuizzesRoot" -ForegroundColor Yellow
            return $result
        }
        $files = @(Get-ChildItem -LiteralPath $QuizzesRoot -Filter '*.txt' -File |
                   Where-Object { $_.Name -notin @('README.txt', 'Prompt.txt', 'GenerateQuizPrompt.txt') -and $_.Name -notlike '~$*' } |
                   Sort-Object LastWriteTime)
    }
    else {
        $item = Get-Item -LiteralPath (Resolve-ExistingPath -Value $Target)
        if ($item.PSIsContainer) {
            $files = @(Get-ChildItem -LiteralPath $item.FullName -Filter '*.txt' -File |
                       Where-Object {
                           $_.Name -notin @('README.txt', 'Prompt.txt', 'GenerateQuizPrompt.txt') -and
                           $_.BaseName -notmatch '_pt\d+$' -and
                           $_.Name -notlike '~$*'
                       } |
                       Sort-Object LastWriteTime)
        }
        else {
            $files = @($item)
        }
    }

    if ($files.Count -eq 0) {
        Write-Host ''
        Write-Host '  No new quiz .txt files are waiting in the Quizzes folder.' -ForegroundColor Yellow
        Write-Host '  Save or drag them directly into:' -ForegroundColor DarkGray
        Write-Host "    $QuizzesRoot" -ForegroundColor DarkGray
        return $result
    }

    Write-Host ''
    Write-Host "  Found $($files.Count) file(s) to sort." -ForegroundColor Cyan

    $needsFixRoot = Join-Path $WorkspaceRoot 'Needs-Fixing'
    $unsortedRoot = Join-Path $QuizzesRoot '_Unsorted'
    $duplicatesRoot = Join-Path $WorkspaceRoot '_backups\duplicates'

    foreach ($file in $files) {
        Write-Host ''
        Write-Host "  File: $($file.Name)" -ForegroundColor Cyan

        $parentFull = $file.DirectoryName.TrimEnd('\', '/')

        # A file already parked in Needs-Fixing is left there.
        if ($parentFull -ieq $needsFixRoot.TrimEnd('\', '/')) {
            Write-Host '    - Already in Needs-Fixing; left as is.' -ForegroundColor DarkGray
            $result.NeedsFixing++
            continue
        }

        # _Unsorted is a manual parking spot already; leave it alone.
        if ($parentFull -ieq $unsortedRoot.TrimEnd('\', '/')) {
            Write-Host '    - Already in _Unsorted; rename it like Subject_WhatQuiz.txt first.' -ForegroundColor DarkGray
            $result.Skipped++
            continue
        }

        $subjectFolder = Resolve-SubjectFolder -FileName $file.Name

        if (-not $subjectFolder) {
            if (-not (Test-Path -LiteralPath $unsortedRoot)) {
                New-Item -ItemType Directory -Path $unsortedRoot -Force | Out-Null
            }
            $destination = Get-AvailableFileName -Folder $unsortedRoot -BaseName $file.BaseName -Extension $file.Extension
            Move-Item -LiteralPath $file.FullName -Destination $destination
            Write-Host "    - No usable subject prefix; moved to: _Unsorted\$([System.IO.Path]::GetFileName($destination))" -ForegroundColor Yellow
            Write-Host '      Rename it like Subject_WhatQuiz.txt, then run Sort again.' -ForegroundColor DarkGray
            $result.Skipped++
            continue
        }

        if (-not (Test-Path -LiteralPath $subjectFolder)) {
            New-Item -ItemType Directory -Path $subjectFolder -Force | Out-Null
            Write-Host "    - Created subject folder: $(Split-Path -Leaf $subjectFolder)" -ForegroundColor DarkGray
        }

        $format = Get-QuizFormatInfo -FilePath $file.FullName
        if ($format.ReadError) {
            Write-Host "    - Could not read the file: $($format.ReadError)" -ForegroundColor Red
            $result.Failed++
            continue
        }

        if ($format.Valid -eq 0) {
            if (-not (Test-Path -LiteralPath $needsFixRoot)) {
                New-Item -ItemType Directory -Path $needsFixRoot -Force | Out-Null
            }
            $destination = Get-AvailableFileName -Folder $needsFixRoot -BaseName $file.BaseName -Extension $file.Extension
            Move-Item -LiteralPath $file.FullName -Destination $destination
            if ($format.Malformed -gt 0) {
                Write-Host "    - Wrong format ($($format.Malformed) 'Question ' line(s) without a colon); moved to Needs-Fixing." -ForegroundColor Yellow
            }
            else {
                Write-Host "    - No 'Question:' lines found; moved to Needs-Fixing." -ForegroundColor Yellow
            }
            $result.NeedsFixing++
            continue
        }

        $subjectFull = $subjectFolder.TrimEnd('\', '/')
        $alreadySorted = ($parentFull -ieq $subjectFull)
        $moveDestination = $file.FullName

        if ($alreadySorted) {
            Write-Host "    - Already in: $([System.IO.Path]::GetFileName($subjectFolder))\" -ForegroundColor DarkGray
        }
        else {
            $candidate = Join-Path $subjectFolder $file.Name
            if (Test-Path -LiteralPath $candidate) {
                $same = $false
                try {
                    $same = ((Get-FileHash -LiteralPath $file.FullName).Hash -eq (Get-FileHash -LiteralPath $candidate).Hash)
                }
                catch {
                    $same = $false
                }

                if ($same) {
                    if (-not (Test-Path -LiteralPath $duplicatesRoot)) {
                        New-Item -ItemType Directory -Path $duplicatesRoot -Force | Out-Null
                    }
                    $duplicateDestination = Get-AvailableFileName -Folder $duplicatesRoot -BaseName $file.BaseName -Extension $file.Extension
                    Move-Item -LiteralPath $file.FullName -Destination $duplicateDestination
                    Write-Host "    - An identical file is already in $([System.IO.Path]::GetFileName($subjectFolder))\; moved the duplicate to _backups\duplicates." -ForegroundColor Yellow
                    $result.Skipped++
                    continue
                }

                $i = 2
                while (Test-Path -LiteralPath $candidate) {
                    $candidate = Join-Path $subjectFolder "$($file.BaseName)($i)$($file.Extension)"
                    $i++
                }
                Write-Host "    - Name already used with different content; saving as: $([System.IO.Path]::GetFileName($candidate))" -ForegroundColor Yellow
            }

            Move-Item -LiteralPath $file.FullName -Destination $candidate
            $moveDestination = $candidate
            Write-Host "    - Moved into: $([System.IO.Path]::GetFileName($subjectFolder))\" -ForegroundColor Green
        }

        try {
            $result.SortedFiles.Add($moveDestination) | Out-Null

            if ($SkipConversion) {
                $result.Sorted++
                Write-Host '    - Filed here; conversion skipped.' -ForegroundColor DarkGray
                continue
            }

            $outFiles = @(Invoke-Conversion -Target $moveDestination)
            $result.Sorted++

            foreach ($outFile in $outFiles) {
                $result.CreatedFiles.Add($outFile) | Out-Null
                $outFolder = Split-Path -Parent $outFile
                if (-not $result.OutputFolders.Contains($outFolder)) {
                    $result.OutputFolders.Add($outFolder) | Out-Null
                }
                Write-Host "    - Excel ready: $([System.IO.Path]::GetFileName($outFile))" -ForegroundColor Green
            }

            if ($outFiles.Count -gt 0) {
                Write-Host "    - Output folder: $($result.OutputFolders[$result.OutputFolders.Count - 1])" -ForegroundColor DarkGray
            }
            else {
                Write-Host '    - Conversion finished, but no new .xlsx file was found.' -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "    - Conversion failed: $($_.Exception.Message)" -ForegroundColor Red
            $result.Failed++
        }
    }

    $uniqueFolders = @($result.OutputFolders | Select-Object -Unique)

    if ($uniqueFolders.Count -gt 0) {
        if (-not $NoOpen) {
            foreach ($folder in $uniqueFolders) {
                try { Start-Process -FilePath $folder } catch { }
            }
        }

        if (-not $NoClipboard) {
            $copyValue = $uniqueFolders[$uniqueFolders.Count - 1]
            if (Set-ClipboardText -Value $copyValue) {
                $result.CopiedPath = $copyValue
                Write-Host ''
                Write-Host "  Copied to clipboard: $copyValue" -ForegroundColor Cyan
            }
            else {
                Write-Host ''
                Write-Host '  Could not copy the output folder path to the clipboard.' -ForegroundColor Yellow
            }
        }
    }

    Write-Host ''
    Write-Host "  Sorted: $($result.Sorted)   Skipped/unplaced: $($result.Skipped)   Needs-Fixing: $($result.NeedsFixing)   Failed: $($result.Failed)" -ForegroundColor Cyan

    $result.Success = ($result.Sorted -gt 0)
    return $result
}

function Invoke-AutoSort {
    <#
        Public wrapper. By default it returns a plain array of the .xlsx
        paths that were just created, ready for Invoke-SmartUpload. Use
        -PassThruResult when the caller also needs the counters (the
        non-menu Sort action uses that for its exit code).
    #>
    param(
        [string]$Target,
        [switch]$SkipConversion,
        [switch]$PassThruResult
    )

    $result = Invoke-AutoSortCore -Target $Target -SkipConversion:$SkipConversion
    if ($PassThruResult) { return $result }
    return @($result.CreatedFiles)
}

# ---------------------------------------------------------------------------
#  Menu
# ---------------------------------------------------------------------------

function Show-Menu {
    param([string]$SelectedPath)

    Clear-Screen
    Write-Banner
    Write-Host ''
    Write-Host '  Turns quiz .txt files into Quizizz Excel files, and can'
    Write-Host '  sort new files by subject, split a quiz, combine parts, or upload'
    Write-Host '  finished workbooks to Wayground.'
    Write-Host ''

    if ($SelectedPath) {
        Write-Host '  Selected: ' -NoNewline
        Write-Host $SelectedPath -ForegroundColor Cyan
    }
    else {
        Write-Host '  Selected: (none)' -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host '  What do you want to do?'
    Write-Host ''
    Write-Host '    [1]  Convert .txt -> Excel   (suggests your newest quiz)'
    Write-Host '    [2]  Sort & convert new quizzes (drop them in Quizzes\)'
    Write-Host '    [3]  Split one quiz into parts'
    Write-Host '    [4]  Combine split parts back into one quiz'
    Write-Host '    [5]  Pick a different file or folder'
    Write-Host '    [6]  Upload existing XLSX to Wayground'
    Write-Host '    [0]  Quit'
    Write-Host ''
    Write-Host '  Tip: save new .txt files directly into the Quizzes folder,'
    Write-Host '       then choose 2. Choose 1 and press Enter for the newest quiz.'
    Write-Host '       Type B in any file prompt to browse with File Explorer.'
    Write-Host '       After converting, the toolkit offers to upload to Wayground.'
    Write-Host ''
}

function Read-ConvertTarget {
    return (Read-QuizTarget -Prompt 'Which quiz? Press Enter for the newest, type a number, B to browse, F for a folder, or drag it here')
}

function Start-Menu {
    param([string]$InitialPath)

    $selectedPath = $null
    if ($InitialPath) {
        try {
            $selectedPath = Resolve-ExistingPath -Value $InitialPath
        }
        catch {
            Write-Host ''
            Write-Host "  The path you dropped could not be used: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host '  Starting with no selection.' -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }

    while ($true) {
        Show-Menu -SelectedPath $selectedPath
        $choice = Read-MenuChoice -Prompt 'Type 1, 2, 3, 4, 5, 6 or 0 and press Enter' -Valid @('1', '2', '3', '4', '5', '6', '0')

        switch ($choice) {
            '1' {
                try {
                    $target = $selectedPath
                    if (-not $target) { $target = Read-ConvertTarget }
                    if (-not $target) { throw 'No quiz was selected.' }

                    $item = Get-Item -LiteralPath $target

                    # A file (or the whole Quizzes folder) sitting in the
                    # general folder is sorted first, then converted.
                    if ($item.PSIsContainer) {
                        if (Test-IsGeneralFolderPath -FolderPath $item.FullName) {
                            $createdFiles = @(Invoke-AutoSort -Target $item.FullName)
                            Invoke-SmartUpload -Files $createdFiles
                        }
                        else {
                            $recurse = Read-YesNo -Prompt 'Also include sub-folders?' -Default $false
                            $createdFiles = @(Invoke-Conversion -Target $target -RecurseSubfolders:$recurse)
                            $outFolder = Get-ConversionOutputFolder -Target $target
                            Write-Host ''
                            if ((Test-Path -LiteralPath $outFolder) -and
                                (Read-YesNo -Prompt 'Open the output folder now?' -Default $true)) {
                                Start-Process -FilePath $outFolder
                            }
                            Invoke-SmartUpload -Files $createdFiles
                        }
                    }
                    else {
                        if (Test-IsGeneralFolderPath -FolderPath $item.DirectoryName) {
                            $createdFiles = @(Invoke-AutoSort -Target $item.FullName)
                            Invoke-SmartUpload -Files $createdFiles
                        }
                        else {
                            $createdFiles = @(Invoke-Conversion -Target $target)
                            $outFolder = Get-ConversionOutputFolder -Target $target
                            Write-Host ''
                            if ((Test-Path -LiteralPath $outFolder) -and
                                (Read-YesNo -Prompt 'Open the output folder now?' -Default $true)) {
                                Start-Process -FilePath $outFolder
                            }
                            Invoke-SmartUpload -Files $createdFiles
                        }
                    }
                }
                catch {
                    Write-Host ''
                    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                }
                Wait-Menu
            }

            '2' {
                try {
                    $target = $selectedPath
                    $createdFiles = @(Invoke-AutoSort -Target $target)
                    Invoke-SmartUpload -Files $createdFiles
                }
                catch {
                    Write-Host ''
                    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                }
                Wait-Menu
            }

            '3' {
                try {
                    $target = $selectedPath
                    if (-not $target) {
                        $target = Read-QuizTarget -Prompt 'Which quiz? Press Enter for the newest, type a number, B to browse, F for a folder, or drag it here'
                    }
                    if (-not $target) { throw 'No quiz was selected.' }

                    $size = Read-Int -Prompt 'How many questions per split file? [Default 10]' -Default 10 -Minimum 1
                    $tail = Read-Int -Prompt 'If the leftover is this many questions or fewer, keep them in one file. [Default 15]' -Default 15 -Minimum 0

                    $split = Invoke-Split -Target $target -Size $size -TailThreshold $tail

                    if ($split.Success) {
                        Write-Host ''
                        if (Read-YesNo -Prompt 'Convert every .txt in that folder to Excel now? (includes the original)' -Default $false) {
                            $createdFiles = New-Object System.Collections.Generic.List[string]
                            $outFolders   = New-Object System.Collections.Generic.List[string]

                            foreach ($splitFolder in @($split.Folders)) {
                                foreach ($outFile in @(Invoke-Conversion -Target $splitFolder)) {
                                    $createdFiles.Add($outFile) | Out-Null
                                }

                                $outFolder = Get-ConversionOutputFolder -Target $splitFolder
                                if (-not $outFolders.Contains($outFolder)) {
                                    $outFolders.Add($outFolder) | Out-Null
                                }
                            }

                            Write-Host ''
                            if ($outFolders.Count -gt 0 -and
                                (Read-YesNo -Prompt 'Open the output folder now?' -Default $true)) {
                                foreach ($outFolder in $outFolders) {
                                    if (Test-Path -LiteralPath $outFolder) {
                                        Start-Process -FilePath $outFolder
                                    }
                                }
                            }
                            Invoke-SmartUpload -Files @($createdFiles)
                        }
                    }
                    else {
                        Write-Host ''
                        Write-Host '  Nothing was split.' -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host ''
                    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                }
                Wait-Menu
            }

            '4' {
                try {
                    $target = $selectedPath
                    if (-not $target) {
                        $target = Read-Path -Prompt 'Drag a split file (_pt1.txt) or its folder here, type B to browse, or type F for a folder, then press Enter:'
                    }

                    $combine = Invoke-Combine -Target $target

                    if ($combine.Success) {
                        Write-Host ''
                        if (Read-YesNo -Prompt 'Open the folder with the combined file(s)?' -Default $true) {
                            Start-Process -FilePath $combine.OutputFolder
                        }
                    }
                    else {
                        Write-Host ''
                        Write-Host '  Nothing was combined.' -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host ''
                    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                }
                Wait-Menu
            }

            '5' {
                $newPath = Read-Path -Prompt 'Drag a file or folder here, type B to browse files, type F for a folder, or press Enter alone to clear' -AllowEmpty

                if ($null -eq $newPath) {
                    $selectedPath = $null
                    Write-Host ''
                    Write-Host '  Selection cleared.' -ForegroundColor DarkGray
                }
                else {
                    $selectedPath = $newPath
                    Write-Host ''
                    Write-Host "  Selected: $selectedPath" -ForegroundColor Cyan
                }
            }

            '6' {
                try {
                    if (Test-WquInstalled) {
                        $defaultUploadFolder = Join-Path $QuizzesRoot 'DataComs\_QuizizzOutput'
                        $uploadFolder = Read-UploadFolder -DefaultFolder $defaultUploadFolder

                        if ($uploadFolder) {
                            $assignArguments = @(Read-AssignArguments)
                            Invoke-WquPublish -Paths @($uploadFolder) -AssignArguments $assignArguments
                        }
                    }
                }
                catch {
                    Write-Host ''
                    Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                }
                Wait-Menu
            }

            '0' {
                Write-Host ''
                Write-Host '  Done.' -ForegroundColor Cyan
                return
            }
        }
    }
}

# ---------------------------------------------------------------------------
#  Main
# ---------------------------------------------------------------------------

if ($QuestionsPerFile -lt 1) {
    Write-Host 'ERROR: -QuestionsPerFile must be 1 or more.' -ForegroundColor Red
    exit 1
}
if ($TailThreshold -lt 0) {
    Write-Host 'ERROR: -TailThreshold must be 0 or more.' -ForegroundColor Red
    exit 1
}

if ($Action -ne 'Menu') {
    try {
        # Sort with no path is valid: it scans the general Quizzes folder.
        if ($Action -ne 'Sort' -and [string]::IsNullOrWhiteSpace($Path)) {
            throw "-Path is required when -Action is $Action."
        }

        $target = $null
        if (-not [string]::IsNullOrWhiteSpace($Path)) {
            $target = Resolve-ExistingPath -Value $Path
        }

        switch ($Action) {
            'Convert' {
                $item = Get-Item -LiteralPath $target
                if ($item.PSIsContainer) {
                    Invoke-Conversion -Target $target -RecurseSubfolders:$Recurse | Out-Null
                }
                else {
                    Invoke-Conversion -Target $target | Out-Null
                }
            }
            'Sort' {
                $r = Invoke-AutoSort -Target $target -PassThruResult
                if ($r.Failed -gt 0) { exit 1 }
            }
            'Split' {
                Invoke-Split -Target $target -Size $QuestionsPerFile -TailThreshold $TailThreshold | Out-Null
            }
            'Upload' {
                if (-not (Test-WquInstalled)) { exit 2 }
                Invoke-WquPublish -Paths @($target) -AssignArguments @('--no-assign')
            }
            'Combine' {
                $r = Invoke-Combine -Target $target
                if (-not $r.Success) { exit 1 }
            }
        }
    }
    catch {
        Write-Host ''
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    exit 0
}

Start-Menu -InitialPath $Path
