# Live Word regression test. Creates a disposable unsaved document only.
$ErrorActionPreference='Stop'
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
    'PASS: native bracket label, numeric range endpoint, dynamic REF renumbering; no file saved.'
} finally {
    if ($null -ne $taskDoc) {$taskDoc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskDoc)}
    if ($null -ne $taskWord) {
        if ($taskWord.Documents.Count -eq 0) {$taskWord.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskWord)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
