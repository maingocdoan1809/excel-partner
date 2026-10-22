param(
  [ValidateRange(0, 600)][int]$WatchSeconds = 0,
  [ValidateRange(200, 5000)][int]$IntervalMs = 500,
  [switch]$SelfTest
)

# Read-only Windows COM experiment. Does not select, edit, save or close anything.
# GetActiveObject attaches to a registered Excel instance; this POC does not
# resolve multiple Excel processes to the add-in's workbook session yet.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

function Read-ShapeSelection {
  param($Application)
  $proxies = [Collections.Generic.List[object]]::new()
  function Keep($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
      $proxies.Add($value)
    }
    return ,$value
  }
  try {
    $excel = if ($null -ne $Application) { $Application } else {
      Keep ([Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'))
    }
    $window = Keep ($excel.ActiveWindow)
    if ($null -eq $window) { return [ordered]@{ status = 'no_window'; shapes = @() } }
    $book = Keep ($window.Parent)
    $sheet = Keep ($window.ActiveSheet)
    $selection = Keep ($window.Selection)
    $result = [ordered]@{
      status = 'ok'; workbook = [string]$book.Name; workbookPath = [string]$book.FullName
      sheet = [string]$sheet.Name; windowHwnd = [string]$window.Hwnd
      selectionKind = 'other'; count = 0; shapes = @()
    }
    $range = $null
    try {
      $range = Keep ($selection.ShapeRange)
    } catch { $range = $null }
    if ($null -ne $range) {
      $result.selectionKind = 'shapes'
      $result.shapes = @(for ($i = 1; $i -le $range.Count; $i++) {
        $shape = Keep ($range.Item($i))
        [ordered]@{ id = [string]$shape.ID; name = [string]$shape.Name
          type = [int]$shape.Type; connector = ($shape.Connector -eq -1) }
      })
      $result.count = $result.shapes.Count
    } else {
      try {
        $address = $selection.Address($false, $false)
        if ($address) { $result.selectionKind = 'cells'; $result.address = [string]$address }
      } catch { }
    }
    return $result
  } catch {
    return [ordered]@{ status = 'unavailable'; error = $_.Exception.Message; shapes = @() }
  } finally {
    for ($i = $proxies.Count - 1; $i -ge 0; $i--) {
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($proxies[$i]) } catch { }
    }
  }
}

if ($SelfTest) {
  # Only this new, hidden application and its unsaved workbook are changed.
  $fixture = $null
  $fixtureBook = $null
  try {
    $fixture = New-Object -ComObject Excel.Application
    $fixture.Visible = $false
    $fixture.DisplayAlerts = $false
    $fixtureBook = $fixture.Workbooks.Add()
    $fixtureSheet = $fixtureBook.Worksheets.Item(1)
    $a = $fixtureSheet.Shapes.AddShape(1, 10, 10, 50, 30); $a.Name = 'POC_A'
    $b = $fixtureSheet.Shapes.AddShape(1, 100, 10, 50, 30); $b.Name = 'POC_B'
    $c = $fixtureSheet.Shapes.AddConnector(1, 60, 25, 100, 25); $c.Name = 'POC_Connector'
    function Assert-Selection([string]$Case, [string[]]$Names) {
      $snapshot = Read-ShapeSelection -Application $fixture
      $actual = @($snapshot.shapes | ForEach-Object { $_.name } | Sort-Object)
      $expected = @($Names | Sort-Object)
      if ($snapshot.status -ne 'ok' -or $snapshot.count -ne $Names.Count -or
          ($actual -join '|') -ne ($expected -join '|')) {
        throw "$Case failed: $(ConvertTo-Json -InputObject $snapshot -Depth 6 -Compress)"
      }
      if ($Names -contains 'POC_Connector') {
        if (-not @($snapshot.shapes | Where-Object { $_.name -eq 'POC_Connector' })[0].connector) {
          throw "$Case failed: connector identity lost"
        }
      }
      [ordered]@{ test = $Case; passed = $true; selection = $snapshot } | ConvertTo-Json -Depth 7 -Compress
    }
    $a.Select($true)
    Assert-Selection 'single shape' @('POC_A')
    $b.Select($false)
    Assert-Selection 'add second shape' @('POC_A', 'POC_B')
    $c.Select($false)
    Assert-Selection 'add connector' @('POC_A', 'POC_B', 'POC_Connector')
    $fixtureSheet.Shapes.Range(@('POC_A', 'POC_Connector')).Select($true)
    Assert-Selection 'remove one from selection' @('POC_A', 'POC_Connector')
    $c.Select($true)
    Assert-Selection 'single connector' @('POC_Connector')
    [void]$fixtureSheet.Range('C5').Select()
    Assert-Selection 'return to cells' @()
  } finally {
    if ($null -ne $fixtureBook) { $fixtureBook.Close($false) }
    if ($null -ne $fixture) { $fixture.Quit() }
    foreach ($proxy in @($a, $b, $c, $fixtureSheet, $fixtureBook, $fixture)) {
      if ($null -ne $proxy -and [Runtime.InteropServices.Marshal]::IsComObject($proxy)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($proxy)
      }
    }
  }
  return
}

$timer = [Diagnostics.Stopwatch]::StartNew()
$previous = $null
do {
  $snapshot = Read-ShapeSelection
  $json = ConvertTo-Json -InputObject $snapshot -Depth 6 -Compress
  if ($json -ne $previous) {
    [ordered]@{ elapsedMs = $timer.ElapsedMilliseconds; selection = $snapshot } |
      ConvertTo-Json -Depth 7 -Compress
    $previous = $json
  }
  if ($timer.Elapsed.TotalSeconds -ge $WatchSeconds) { break }
  Start-Sleep -Milliseconds $IntervalMs
} while ($true)
