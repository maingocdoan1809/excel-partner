# Shared by the bootstrap and the launcher embedded in the MSI. It must work before any package files exist.
$ErrorActionPreference = 'Stop'
if (-not $dataDir) { $dataDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Excel Partner' }
if (-not $logFile) {
  $logFile = Join-Path $dataDir 'msi-install.log'
  try { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null } catch { $logFile = Join-Path ([IO.Path]::GetTempPath()) 'office-partner-msi-install.log' }
}

function Write-Log([string] $message) {
  $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'), $message
  try { Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8 } catch { }
  Write-Host $message
}

function Stop-ProductProcesses {
  $markers = @($dataDir, 'Excel Partner', 'office-partner', 'office-partner-node', 'agent-gateway', 'recovery.mjs', 'dist\bridge\src\index.js', 'dist/bridge/src/index.js')
  $stopped = 0
  try {
    foreach ($process in (Get-CimInstance Win32_Process -Filter "Name='node.exe'" -ErrorAction Stop)) {
      $command = "$($process.CommandLine)"
      $hit = $false
      foreach ($marker in $markers) { if ($command.ToLower().Contains($marker.ToLower())) { $hit = $true; break } }
      if (-not $hit) { continue }
      Write-Log "    stopping node $($process.ProcessId): $command"
      try { Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop; $stopped++ } catch { Write-Log "    could not stop $($process.ProcessId): $($_.Exception.Message)" }
    }
  } catch { Write-Log "    process list unavailable: $($_.Exception.Message)" }
  # A bridge started by another copy (or by an older installer) may answer on its port without matching a name.
  foreach ($port in 8765, 8766) {
    try {
      $owners = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction Stop | Select-Object -ExpandProperty OwningProcess -Unique
      foreach ($owner in $owners) {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId=$owner" -ErrorAction SilentlyContinue
        if ($process -and "$($process.Name)" -match '^node') {
          Write-Log "    stopping node $owner that holds port $port"
          Stop-Process -Id $owner -Force -ErrorAction SilentlyContinue
          $stopped++
        } elseif ($process) { Write-Log "    port $port is held by $($process.Name) ($owner), left alone" }
      }
    } catch { }
  }
  Write-Log "    stopped $stopped process(es)"
  Start-Sleep -Milliseconds 700
}
