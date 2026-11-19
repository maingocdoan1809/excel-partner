param([string] $Root, [string] $Work, [string] $NodeExe)
$ErrorActionPreference = 'Stop'

# Load only the functions under test. Never execute installer actions or read the user's registry and data.
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $Root 'installer\msi\bootstrap.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors -join "`n") }
foreach ($name in 'Ensure-Node', 'Ensure-NpmLayout', 'Remove-CertificateTrust') {
  $definition = $ast.Find({ param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name }, $true)
  . ([scriptblock]::Create($definition.Extent.Text))
}
function Write-Log([string] $message) { Write-Output $message | Out-Host }

$Architecture = 'x64'
$nodeHome = Join-Path $Work 'local\Programs\nodejs'
$altNodeHome = Join-Path $Work 'local\Programs\office-partner-node'
New-Item -ItemType Directory -Path $nodeHome, $altNodeHome -Force | Out-Null
Set-Content -LiteralPath (Join-Path $nodeHome 'node.exe') -Value ''
Set-Content -LiteralPath (Join-Path $altNodeHome 'node.exe') -Value ''
function Get-NodeVersion { return 'fixture' }
function Get-NodeRuntime([string] $exe) {
  if ($exe -eq (Join-Path $altNodeHome 'node.exe')) { return 'x64:127' }
  return $script:foreignRuntime
}
foreach ($foreignRuntime in 'x64:137', 'arm64:127') {
  $found = Ensure-Node ([pscustomobject]@{ node = [pscustomobject]@{ abi = 127 } }) -NoDownload
  if ($found -ne (Join-Path $altNodeHome 'node.exe')) { throw "Selected an incompatible Node.js: $foreignRuntime" }
}

$Package = Join-Path $Work 'prefix\node_modules\office-partner'
$npmRoot = Join-Path $Work 'other-prefix\node_modules'
New-Item -ItemType Directory -Path $Package -Force | Out-Null
Set-Content -LiteralPath (Join-Path $Work 'prefix\office-partner.cmd') -Value 'fixture shim'
function Get-NpmRoot { return $npmRoot }
Ensure-NpmLayout 'fixture-node'
if ((Get-Content -LiteralPath (Join-Path $Work 'other-prefix\office-partner.cmd') -Raw).Trim() -ne 'fixture shim') { throw 'The command shim was not copied to the npm prefix.' }
if (-not ((Get-Item -LiteralPath (Join-Path $npmRoot 'office-partner')).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'The alternate npm prefix has no package junction.' }

$gatewayDir = Join-Path $Work 'gateway'
New-Item -ItemType Directory -Path $gatewayDir -Force | Out-Null
$gatewayCert = Join-Path $gatewayDir 'localhost.crt'
Copy-Item -LiteralPath (Join-Path $Work 'certificate.crt') -Destination $gatewayCert
$thumbprint = (New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $gatewayCert).Thumbprint
$untrusted = $false
function certutil.exe { $script:untrusted = $true; $global:LASTEXITCODE = 0 }
Remove-CertificateTrust $thumbprint
if ($untrusted) { throw 'Removing the add-in also untrusted the gateway certificate.' }
Remove-Item -LiteralPath $gatewayCert
Remove-CertificateTrust $thumbprint
if (-not $untrusted) { throw 'A certificate with no remaining gateway owner was not untrusted.' }

# Exercise the generated .cmd with an incompatible primary Node and a real, compatible alternate Node.
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$source = Join-Path $Work 'incompatible-node.cs'
Set-Content -LiteralPath $source -Value 'class IncompatibleNode { static int Main() { return 1; } }'
& $compiler /nologo /target:exe "/out:$(Join-Path $nodeHome 'node.exe')" $source
if ($LASTEXITCODE -ne 0) { throw 'The incompatible Node fixture did not compile.' }
Copy-Item -LiteralPath $NodeExe -Destination (Join-Path $altNodeHome 'node.exe') -Force
$fixtureCli = Join-Path $Package 'bin\office-partner.mjs'
New-Item -ItemType Directory -Path (Split-Path -Parent $fixtureCli) -Force | Out-Null
Set-Content -LiteralPath $fixtureCli -Value 'console.log(JSON.stringify(process.argv.slice(2))); process.exit(17);'
Copy-Item -LiteralPath (Join-Path $Work 'shim.cmd') -Destination (Join-Path $Work 'prefix\office-partner.cmd') -Force
$env:LOCALAPPDATA = Join-Path $Work 'local'
$result = & (Join-Path $Work 'prefix\office-partner.cmd') 'hello world' 'status'
if ($LASTEXITCODE -ne 17) { throw "The command shim lost the CLI exit code: $LASTEXITCODE" }
if (($result | ConvertFrom-Json) -join ',' -ne 'hello world,status') { throw 'The command shim lost or split CLI arguments.' }
Write-Output 'PASS: bootstrap selection, npm layout, shared trust and command shim.'
