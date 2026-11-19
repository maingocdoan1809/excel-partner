# Embedded in the MSI command so cleanup can run after the installed files are removed.
function Remove-ProductDataTree([IO.DirectoryInfo] $Directory) {
  foreach ($entry in $Directory.EnumerateFileSystemInfos()) {
    if ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) {
      $entry.Delete()
    } elseif ($entry -is [IO.DirectoryInfo]) {
      Remove-ProductDataTree $entry
    } else {
      if ($entry.IsReadOnly) { $entry.IsReadOnly = $false }
      $entry.Delete()
    }
  }
  $Directory.Delete()
}

function Remove-ProductData {
  param(
    [switch] $Gateway,
    [string] $LocalAppData = [Environment]::GetFolderPath('LocalApplicationData'),
    [string] $UserProfile = [Environment]::GetFolderPath('UserProfile')
  )

  if ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem) { throw 'Run the uninstall as the user who installed Office Partner.' }

  $folders = @(
    @{ Root = $LocalAppData; Name = 'Excel Partner' },
    @{ Root = $UserProfile; Name = '.excel-partner' }
  )
  if ($Gateway) { $folders += @{ Root = $LocalAppData; Name = 'Agent Gateway' } }

  $failures = @()
  foreach ($folder in $folders) {
    try {
      if (-not $folder.Root) { throw 'Windows did not return a user folder.' }
      $target = [IO.DirectoryInfo] (Join-Path $folder.Root $folder.Name)
      if (-not $target.Exists) { continue }
      for ($directory = $target; $directory; $directory = $directory.Parent) {
        if ($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) {
          throw "Refusing to follow a redirected folder: $($directory.FullName)"
        }
      }
      Remove-ProductDataTree $target
    } catch {
      $failures += "$($folder.Name): $($_.Exception.Message)"
    }
  }
  if ($failures.Count) {
    $message = $failures -join [Environment]::NewLine
    $log = Join-Path ([IO.Path]::GetTempPath()) 'office-partner-msi-purge.log'
    try { Add-Content -LiteralPath $log -Value $message -Encoding UTF8 -ErrorAction Stop } catch { }
    throw $message
  }
}
