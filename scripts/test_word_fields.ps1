# Live Word regression test. Only synthetic fixtures; never opens a user's document.
[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path ([IO.Path]::GetTempPath()) ('academic-word-test-' + [guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Test output directory must be new.' }
$taskQa = (New-Item -ItemType Directory -Path $OutputDirectory).FullName
$taskInput = Join-Path $taskQa 'source.docx'
$taskOutput = Join-Path $taskQa 'final.docx'
$taskWord=$null
$taskDoc=$null
try {
    $taskWord=New-Object -ComObject Word.Application
    $taskWord.Visible=$false
    $taskWord.DisplayAlerts=0
    $taskWord.AutomationSecurity=3
    $taskDoc=$taskWord.Documents.Add()
    $taskDoc.Content.Text="First reference`rSecond reference`rCitation test`r"
    $taskTemplate=$taskDoc.ListTemplates.Add($false,'AcademicWordTest')
    $taskLevel=$taskTemplate.ListLevels.Item(1)
    $taskLevel.NumberFormat='[%1]'
    $taskLevel.NumberStyle=0
    $taskLevel.StartAt=1
    $taskLevel.TrailingCharacter=0
    $taskRange=$taskDoc.Range($taskDoc.Paragraphs.Item(1).Range.Start,$taskDoc.Paragraphs.Item(2).Range.End)
    $taskRange.ListFormat.ApplyListTemplateWithLevel($taskTemplate,$false,0,2,1)
    [void]$taskDoc.Bookmarks.Add('AWBibTest',$taskDoc.Paragraphs.Item(2).Range)
    $taskFields=@()
    foreach ($taskCode in @('REF AWBibTest \n \h','REF AWBibTest \n \h \# "0"')) {
        $taskRange=$taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1)
        $taskField=$taskDoc.Fields.Add($taskRange,-1,$taskCode,$false)
        [void]$taskField.Update()
        $taskFields += $taskField
        $taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1).InsertAfter("`r")
    }
    if ($taskFields[0].Result.Text -ne '[2]' -or $taskFields[1].Result.Text -ne '2') { throw 'Initial native numbering failed' }
    # Change the native sequence; existing REF fields must follow their bookmarked target.
    $taskLevel.StartAt=5
    foreach ($taskField in $taskFields) { [void]$taskField.Update() }
    if ($taskFields[0].Result.Text -ne '[6]' -or $taskFields[1].Result.Text -ne '6') { throw 'Dynamic REF renumbering failed' }
    # Caption bookmark deliberately uses neither AWFig nor _RefFig.
    $taskStart = $taskDoc.Content.End-1
    $taskDoc.Range($taskStart,$taskStart).InsertAfter('图 ')
    $taskRange=$taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1)
    $taskCaption=$taskDoc.Fields.Add($taskRange,-1,'SEQ 图 \* ARABIC',$false)
    [void]$taskCaption.Update()
    $taskCaption.Result.Text='99'
    [void]$taskDoc.Bookmarks.Add('_RefCalibrationTest',$taskDoc.Range($taskStart,$taskCaption.Result.End+1))
    $taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1).InsertAfter(" Test caption`r")
    $taskRange=$taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1)
    $taskRef=$taskDoc.Fields.Add($taskRange,-1,'REF _RefCalibrationTest \h',$false)
    [void]$taskRef.Update()
    if ($taskRef.Result.Text -ne '图 99') { throw 'Synthetic caption bookmark range incorrect.' }
    $taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1).InsertAfter("`r")
    $taskRange=$taskDoc.Range($taskDoc.Content.End-1,$taskDoc.Content.End-1)
    $taskEquation=$taskDoc.Fields.Add($taskRange,-1,'SEQ MTEqn',$false)
    $taskEquation.Result.Text='42'
    $taskEquation.Locked=$true
    $taskDoc.SaveAs2([string]$taskInput,16)
} finally {
    if ($null -ne $taskDoc) {$taskDoc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskDoc)}
    if ($null -ne $taskWord) {
        if ($taskWord.Documents.Count -eq 0) {$taskWord.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskWord)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
# Use a non-default size so the saved XML must carry an explicit override.
$taskReport = & (Join-Path $PSScriptRoot 'word_finalize.ps1') -InputPath $taskInput -OutputPath $taskOutput -UpdateFigures -FigureReferenceSize 13.5 | ConvertFrom-Json
if ($taskReport.sequences -ne 1 -or $taskReport.references.Count -ne 1 -or $taskReport.references[0].result -ne '图 1' -or -not $taskReport.input_unchanged) {
    throw 'Selective caption update failed.'
}
# Inspect the saved package without opening a second Word instance.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$taskZip=[IO.Compression.ZipFile]::OpenRead($taskOutput)
try {
    $taskReader=[IO.StreamReader]::new($taskZip.GetEntry('word/document.xml').Open())
    try { [xml]$taskXml=$taskReader.ReadToEnd() } finally { $taskReader.Dispose() }
    $taskNs=[Xml.XmlNamespaceManager]::new($taskXml.NameTable)
    $taskNs.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $taskEquationCode=$taskXml.SelectSingleNode('//w:instrText[contains(., "SEQ MTEqn")]|//w:fldSimple[contains(@w:instr, "SEQ MTEqn")]', $taskNs)
    if ($null -eq $taskEquationCode) { throw 'Protected equation field lost.' }
    if ($taskEquationCode.LocalName -eq 'fldSimple') {
        $taskEquationResult=$taskEquationCode.SelectSingleNode('.//w:t', $taskNs).InnerText
        $taskLocked=$taskEquationCode.GetAttribute('fldLock',$taskNs.LookupNamespace('w'))
    } else {
        $taskEquationParagraph=$taskEquationCode.ParentNode.ParentNode
        $taskEquationResult=$taskEquationParagraph.SelectSingleNode('.//w:t', $taskNs).InnerText
        $taskLocked=$taskEquationParagraph.SelectSingleNode('.//w:fldChar[@w:fldCharType="begin"]', $taskNs).GetAttribute('fldLock',$taskNs.LookupNamespace('w'))
    }
    if ($taskEquationResult -ne '42' -or $taskLocked -notin @('1','true','on')) { throw 'Protected equation field changed.' }
    $taskRefCode=$taskXml.SelectSingleNode('//w:instrText[contains(., "REF _RefCalibrationTest")]|//w:fldSimple[contains(@w:instr, "REF _RefCalibrationTest")]', $taskNs)
    if ($taskRefCode.LocalName -eq 'fldSimple') { $taskRefScope=$taskRefCode } else { $taskRefScope=$taskRefCode.ParentNode.ParentNode }
    $taskSize=$taskRefScope.SelectSingleNode('.//w:r[w:t]/w:rPr/w:sz', $taskNs)
    if ($null -eq $taskSize -or $taskSize.GetAttribute('val',$taskNs.LookupNamespace('w')) -ne '27') { throw 'Body cross-reference size not applied.' }
} finally { $taskZip.Dispose() }
"PASS: native bibliography numbering, numeric range endpoints, caption REF update, body font size, protected equation field, unchanged input. Synthetic fixtures: $taskQa"
