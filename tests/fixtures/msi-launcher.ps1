param([string] $Root, [string] $Work)
$ErrorActionPreference = 'Stop'
$exe = Join-Path $Work 'run-hidden.exe'
$resource = Join-Path $Root 'installer\msi\stop-processes.ps1'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /optimize+ "/out:$exe" "/resource:${resource},StopProcesses" (Join-Path $Root 'installer\msi\run-hidden.cs')
if ($LASTEXITCODE -ne 0) { throw 'The MSI launcher did not compile.' }

$assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($exe))
$method = $assembly.GetType('RunHidden').GetMethod('StopArguments', [Reflection.BindingFlags]'NonPublic,Static')
$arguments = $method.Invoke($null, @())
$encoded = ($arguments -split ' ')[-1]
$script = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($encoded))
if ($script -notmatch 'Stop-ProductProcesses\s*$') { throw 'The launcher does not invoke the embedded stop function.' }
foreach ($file in @('bootstrap.ps1', 'stop-processes.ps1')) {
  $tokens = $null; $errors = $null
  [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $Root "installer\msi\$file"), [ref]$tokens, [ref]$errors)
  if ($errors.Count) { throw ($errors -join "`n") }
}

# Every process operation is mocked. No installed product, user data, or real process is accessed.
$dataDir = Join-Path $Work 'data'
$logFile = Join-Path $Work 'stop.log'
$mockStops = New-Object System.Collections.Generic.List[int]
function Get-CimInstance {
  param($ClassName, [string] $Filter)
  if ($Filter -like 'Name=*') {
    return @(
      [pscustomobject]@{ ProcessId = 101; Name = 'node.exe'; CommandLine = 'node C:\npm\node_modules\office-partner\dist\bridge\src\index.js' },
      [pscustomobject]@{ ProcessId = 202; Name = 'node.exe'; CommandLine = 'node C:\other\server.js' }
    )
  }
  if ($Filter -eq 'ProcessId=303') { return [pscustomobject]@{ Name = 'node.exe' } }
  if ($Filter -eq 'ProcessId=404') { return [pscustomobject]@{ Name = 'other.exe' } }
  throw "Unexpected process query: $Filter"
}
function Get-NetTCPConnection {
  param($State, [int] $LocalPort)
  return [pscustomobject]@{ OwningProcess = $(if ($LocalPort -eq 8765) { 303 } else { 404 }) }
}
function Stop-Process { param([int] $Id, [switch] $Force) $mockStops.Add($Id) }
function Start-Sleep { param([int] $Milliseconds) }
& ([scriptblock]::Create($script))
if (($mockStops -join ',') -ne '101,303') { throw "Unexpected processes stopped: $($mockStops -join ',')" }
Write-Output 'PASS: embedded stop works without an installed bootstrap and leaves unrelated processes alone.'
