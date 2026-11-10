# Validates Power Query M syntax using the Mashup parser that ships with Excel.
# Read-only: it only loads the parser assemblies and parses text.
$ErrorActionPreference = 'Stop'

$bin    = "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin"
$docAsm = [System.Reflection.Assembly]::LoadFrom((Join-Path $bin 'Microsoft.Mashup.Document.dll'))
$engAsm = [System.Reflection.Assembly]::LoadFrom((Join-Path $bin 'Microsoft.MashupEngine.dll'))

$eng = $engAsm.GetType('Microsoft.Mashup.Engine1.Engine').
    GetProperty('Instance', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static).
    GetValue($null)

$edT     = $docAsm.GetType('Microsoft.Mashup.Host.Document.ExpressionDocument')
$newM    = $edT.GetMethods() | Where-Object { $_.Name -eq 'New' } | Select-Object -First 1
$getText = $edT.GetMethods() | Where-Object { $_.Name -eq 'GetExpressionText' -and $_.GetParameters().Count -eq 1 } | Select-Object -First 1

function Test-MCode {
    param([string]$Label, [string]$Code)
    try {
        $doc  = $newM.Invoke($null, @($eng, $Code))
        $args = New-Object object[] 1
        $null = $getText.Invoke($doc, $args)
        $err  = $args[0]
        if ($null -ne $err) {
            Write-Host ("  FAIL  {0}`n        {1}" -f $Label, $err.ToString())
            return $false
        }
        Write-Host ("  OK    {0}" -f $Label)
        return $true
    }
    catch {
        $inner = $_.Exception.InnerException
        $msg   = if ($inner) { $inner.Message } else { $_.Exception.Message }
        Write-Host ("  THROW {0}`n        {1}" -f $Label, $msg)
        return $false
    }
}

Write-Host '=== parser self-test (must FAIL for the bad ones) ==='
Test-MCode 'valid:   let x = 1 in x' 'let x = 1 in x'            | Out-Null
Test-MCode 'invalid: let x = 1 in'   'let x = 1 in'              | Out-Null
Test-MCode 'invalid: let x = in x'   'let x = in x'              | Out-Null
Test-MCode 'invalid: unbalanced'     'let x = (1 in x'           | Out-Null

Write-Host ''
Write-Host '=== the real query files ==='
$pqDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'powerquery'
$all = $true
foreach ($f in (Get-ChildItem $pqDir -Filter '*.pq' | Sort-Object Name)) {
    $code = [System.IO.File]::ReadAllText($f.FullName)
    $ok = Test-MCode $f.Name $code
    if (-not $ok) { $all = $false }
}

Write-Host ''
if ($all) { Write-Host 'RESULT: all query files parsed cleanly.' }
else      { Write-Host 'RESULT: at least one query file has a syntax problem.' }
