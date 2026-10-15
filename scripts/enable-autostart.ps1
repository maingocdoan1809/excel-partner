<#
  Wake WSL and its bridge user service when this Windows user signs in. Run from WSL with:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File <Windows path to this script> -Distro Ubuntu -WslRoot /home/user/excel-ai-addin
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$Distro,
  [Parameter(Mandatory = $true)] [string]$WslRoot
)

$ErrorActionPreference = 'Stop'
$wslExe = Join-Path $env:WINDIR 'System32\wsl.exe'
$startup = [Environment]::GetFolderPath('Startup')
if (-not $startup) { throw 'Windows Startup folder is unavailable.' }
$shortcutPath = Join-Path $startup 'Excel Partner Bridge.lnk'

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $wslExe
$shortcut.Arguments = '-d "{0}" --cd "{1}" --exec systemctl --user start excel-partner-bridge.service' -f $Distro, $WslRoot
$shortcut.WorkingDirectory = Split-Path $wslExe
$shortcut.WindowStyle = 7
$shortcut.Description = 'Start Excel Partner bridge in WSL at Windows sign-in'
$shortcut.Save()
if (-not (Test-Path $shortcutPath)) { throw 'Could not create the Startup shortcut.' }

Start-Process -FilePath $shortcutPath

Write-Host "Autostart enabled: $shortcutPath"
Write-Host 'Windows will launch the bridge at user logon. Started it once now.'
