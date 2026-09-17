<#
Refresh only requested native Word fields, save a NEW document and/or export PDF.
Never save the input, unlock fields, update ADDIN/MathType fields, or close a user's document.
Example: .\word_finalize.ps1 -InputPath draft.docx -OutputPath final.docx -PdfPath qa.pdf -UpdateBibliography
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [string]$OutputPath,
    [string]$PdfPath,
    [string]$ReportPath,
    [switch]$UpdateBibliography,
    [switch]$UpdateFigures
)
$ErrorActionPreference = 'Stop'
$taskInput = (Resolve-Path -LiteralPath $InputPath).Path
$taskHashBefore = (Get-FileHash -LiteralPath $taskInput -Algorithm SHA256).Hash
if (($UpdateBibliography -or $UpdateFigures) -and -not $OutputPath) {
    throw 'Field updates require a new OutputPath; input is never saved.'
}
$taskResolvedTargets = @()
foreach ($taskTarget in @($OutputPath, $PdfPath, $ReportPath)) {
    if (-not $taskTarget) { continue }
    $taskAbsolute = [IO.Path]::GetFullPath($taskTarget)
    if ($taskResolvedTargets -contains $taskAbsolute) { throw 'Output, PDF, and report paths must be different.' }
    $taskResolvedTargets += $taskAbsolute
    if ($taskAbsolute -eq $taskInput -or (Test-Path -LiteralPath $taskAbsolute)) {
        throw "Refusing to overwrite: $taskAbsolute"
    }
    if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($taskAbsolute)))) {
        throw "Create the output directory first: $taskAbsolute"
    }
}
$taskWord = $null
$taskDoc = $null
$taskReport = [ordered]@{source=$taskInput; source_sha256=$taskHashBefore; sequences=0; references=@(); numbered_references=@()}
try {
    $taskWord = New-Object -ComObject Word.Application
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskWord.AutomationSecurity = 3
    # Open read-only. SaveAs creates a separate deliverable without touching the input.
    $taskDoc = $taskWord.Documents.Open($taskInput, $false, $true, $false)
    if ($taskDoc.Revisions.Count -gt 0 -and ($UpdateBibliography -or $UpdateFigures)) {
        throw 'Tracked revisions present; obtain disposition first.'
    }
    if ($UpdateFigures) {
        foreach ($taskField in $taskDoc.Fields) {
            if ($taskField.Code.Text.Trim() -match '^SEQ\s+(图|表)(?:\s|$)') {
                if ($taskField.Locked) { throw 'A requested caption field is locked; inspect before changing.' }
                [void]$taskField.Update()
                $taskReport.sequences++
            }
        }
    }
    if ($UpdateBibliography) {
        foreach ($taskBookmark in $taskDoc.Bookmarks) {
            if ($taskBookmark.Name -match '^AWBib(\d+)$') {
                $taskList = $taskBookmark.Range.ListFormat
                $taskLabel = $taskList.ListString
                if ($taskLabel -notmatch '^\[\d+\]$') { throw "Not native [n] numbering: $($taskBookmark.Name) $taskLabel" }
                $taskReport.numbered_references += @{bookmark=$taskBookmark.Name; label=$taskLabel; listType=$taskList.ListType}
            }
        }
        if ($taskReport.numbered_references.Count -eq 0) { throw 'No AWBib bookmarks; use docx_ops.py bibliography first.' }
    }
    foreach ($taskField in $taskDoc.Fields) {
        $taskCode = $taskField.Code.Text.Trim()
        $taskIsBib = $UpdateBibliography -and $taskCode -match '^REF\s+(AWBib\d+)\s'
        $taskIsFig = $UpdateFigures -and $taskCode -match '^REF\s+(AWFig\d+|_RefFig\d+)\s'
        if (-not ($taskIsBib -or $taskIsFig)) { continue }
        if ($taskField.Locked) { throw "Requested field locked: $taskCode" }
        $taskBookmarkName = [regex]::Match($taskCode, '^REF\s+(\S+)').Groups[1].Value
        if (-not $taskDoc.Bookmarks.Exists($taskBookmarkName)) { throw "Missing bookmark: $taskBookmarkName" }
        [void]$taskField.Update()
        $taskResult = $taskField.Result.Text
        if ($taskIsBib) {
            $taskExpected = $taskDoc.Bookmarks.Item($taskBookmarkName).Range.ListFormat.ListString
            if ($taskCode -match '\\#\s+"0"') { $taskExpected = $taskExpected.Trim('[', ']') }
        } else {
            $taskExpected = $taskDoc.Bookmarks.Item($taskBookmarkName).Range.Text
            if ($taskExpected -notmatch '^(图|表)\d+$') { throw "Caption bookmark is not label and number only: $taskExpected" }
        }
        if ($taskResult -ne $taskExpected) { throw "Field result mismatch: $taskCode => $taskResult; expected $taskExpected" }
        $taskReport.references += @{code=$taskCode; result=$taskResult}
    }
    $taskDoc.Repaginate()
    $taskReport.pages = $taskDoc.ComputeStatistics(2)
    if ($OutputPath) {
        $taskOutput = [IO.Path]::GetFullPath($OutputPath)
        $taskDoc.SaveAs2($taskOutput, 16)
        $taskReport.output = $taskOutput
    }
    if ($PdfPath) {
        $taskPdf = [IO.Path]::GetFullPath($PdfPath)
        $taskDoc.ExportAsFixedFormat($taskPdf, 17)
        $taskReport.pdf = $taskPdf
    }
} finally {
    if ($null -ne $taskDoc) {
        $taskDoc.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskDoc)
    }
    if ($null -ne $taskWord) {
        if ($taskWord.Documents.Count -eq 0) { $taskWord.Quit() }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskWord)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ((Get-FileHash -LiteralPath $taskInput -Algorithm SHA256).Hash -ne $taskHashBefore) { throw 'Input changed during processing; do not deliver without reconciling.' }
$taskReport.input_unchanged = $true
$taskJson = $taskReport | ConvertTo-Json -Depth 6
if ($ReportPath) { $taskJson | Set-Content -LiteralPath $ReportPath -Encoding UTF8 }
$taskJson
