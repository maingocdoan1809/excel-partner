$ErrorActionPreference = 'Stop'
$probeDir = Join-Path $PSScriptRoot ('../run/rollback-probe-' + [Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($probeDir)
$app = $null
$book = $null
$copy = $null
try {
  # A new hidden Excel instance; never attach to the user's open workbooks.
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $app.EnableEvents = $false
  $app.AutomationSecurity = 3
  $book = $app.Workbooks.Add()
  $sheet = $book.Worksheets.Item(1)
  $sheet.Name = 'Probe'
  $sheet.Range('A1').Value2 = 'Input'
  $sheet.Range('B1').Value2 = 'Result'
  $sheet.Range('A2').Value2 = 7
  $sheet.Range('A3').Value2 = 9
  $sheet.Range('B2').Formula = '=A2*2'
  $sheet.Range('B3').Formula = '=SUM(A2:A3)'
  $sheet.Range('A2').Interior.Color = 123456
  $sheet.Range('A2').Font.Bold = $true
  $sheet.Range('D1:E1').Merge()
  $sheet.Range('D1').Value2 = 'Merged'
  [void]$book.Names.Add('ProbeInput', '=Probe!$A$2')
  [void]$sheet.ListObjects.Add(1, $sheet.Range('A1:B3'), $null, 1)
  $shape = $sheet.Shapes.AddShape(1, 100, 100, 80, 30)
  $shape.Name = 'ProbeShape'
  $shape.TextFrame2.TextRange.Text = 'Snapshot'
  $chart = $sheet.ChartObjects().Add(200, 100, 200, 120)
  $chart.Chart.SetSourceData($sheet.Range('A1:B3'))
  $results = @()
  foreach ($scenario in @('unsaved', 'saved-with-unsaved-edits')) {
    if ($scenario -ne 'unsaved') {
      $book.SaveAs((Join-Path $probeDir 'original.xlsx'), 51)
      $sheet.Range('A2').Value2 = 11
    }
    $expectedValue = $sheet.Range('A2').Value2
    $before = @{ Name = $book.Name; Path = $book.Path; Saved = $book.Saved }
    $snapshot = Join-Path $probeDir ($scenario + '.xlsx')
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $book.SaveCopyAs($snapshot)
    $watch.Stop()
    $unchanged = $book.Name -eq $before.Name -and $book.Path -eq $before.Path -and $book.Saved -eq $before.Saved
    $sheet.Range('A2').Value2 = 99
    $copy = $app.Workbooks.Open($snapshot, 0, $true)
    $restored = $copy.Worksheets.Item('Probe')
    $checks = [ordered]@{
      includesUnsavedValue = $restored.Range('A2').Value2 -eq $expectedValue
      formula = $restored.Range('B2').Formula -eq '=A2*2'
      fill = $restored.Range('A2').Interior.Color -eq 123456
      bold = $restored.Range('A2').Font.Bold -eq $true
      mergedRange = $restored.Range('D1').MergeArea.Address() -eq '$D$1:$E$1'
      table = $restored.ListObjects.Count -eq 1
      shape = $restored.Shapes.Item('ProbeShape').TextFrame2.TextRange.Text -eq 'Snapshot'
      chart = $restored.ChartObjects().Count -eq 1
      namedRange = $copy.Names.Item('ProbeInput').RefersTo -eq '=Probe!$A$2'
      sourceIdentityAndSavedFlagUnchanged = $unchanged
    }
    $copy.Close($false)
    $copy = $null
    $results += [pscustomobject]@{ scenario = $scenario; elapsedMs = $watch.ElapsedMilliseconds; checks = $checks }
    if ($checks.Values -contains $false) { throw ('Snapshot checks failed: ' + ($checks | ConvertTo-Json -Compress)) }
    $sheet.Range('A2').Value2 = 7
  }
  [pscustomobject]@{ excelVersion = $app.Version; excelBuild = $app.Build; results = $results } | ConvertTo-Json -Depth 6
} finally {
  if ($copy) { $copy.Close($false) }
  if ($book) { $book.Close($false) }
  if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
