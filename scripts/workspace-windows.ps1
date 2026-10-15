param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('pick', 'workbook', 'server')]
  [string]$Mode,
  [string]$WorkbookName = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Get-WorkbookPath([string]$Name) {
  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    $book = $null
    if ($Name) {
      foreach ($candidate in $excel.Workbooks) {
        if ($candidate.Name -eq $Name) { $book = $candidate; break }
      }
    } else {
      $book = $excel.ActiveWorkbook
    }
    if ($book -and $book.Path) { return [string]$book.FullName }
  } catch {
    # Excel is not running, the workbook is unsaved, or COM is unavailable.
  }
  return $null
}

function Show-WorkspacePicker {
  return [WorkspacePicker.Dialog]::Choose()
}

# Windows PowerShell runs on .NET Framework, whose FolderBrowserDialog still uses
# the old tree view. IFileDialog with FOS_PICKFOLDERS is the Explorer folder picker.
$pickerSource = @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WorkspacePicker {
  [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
  public class FileOpenDialogCom { }

  [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IFileDialog {
    [PreserveSig] int Show(IntPtr owner);
    void SetFileTypes(uint count, IntPtr specs);
    void SetFileTypeIndex(uint index);
    void GetFileTypeIndex(out uint index);
    void Advise(IntPtr events, out uint cookie);
    void Unadvise(uint cookie);
    void SetOptions(uint options);
    void GetOptions(out uint options);
    void SetDefaultFolder(IntPtr folder);
    void SetFolder(IntPtr folder);
    void GetFolder(out IntPtr folder);
    void GetCurrentSelection(out IntPtr item);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name);
    void GetFileName(out IntPtr name);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
    void GetResult(out IShellItem item);
    void AddPlace(IntPtr item, int location);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
    void Close(int result);
    void SetClientGuid(ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr filter);
  }

  [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IShellItem {
    void BindToHandler(IntPtr context, ref Guid handler, ref Guid iid, out IntPtr result);
    void GetParent(out IShellItem parent);
    void GetDisplayName(uint nameType, [MarshalAs(UnmanagedType.LPWStr)] out string name);
    void GetAttributes(uint mask, out uint attributes);
    void Compare(IShellItem other, uint hint, out int order);
  }

  public static class Dialog {
    public static string Choose() {
      var dialog = (IFileDialog)new FileOpenDialogCom();
      var owner = new Form { TopMost = true, ShowInTaskbar = false, Opacity = 0 };
      try {
        uint options;
        dialog.GetOptions(out options);
        dialog.SetOptions(options | 0x20u | 0x40u); // FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM
        dialog.SetTitle("Add workspace · Excel Partner");
        owner.Show();
        int result = dialog.Show(owner.Handle);
        if (result == unchecked((int)0x800704C7)) return null; // user cancelled
        if (result != 0) Marshal.ThrowExceptionForHR(result);
        IShellItem item;
        dialog.GetResult(out item);
        try {
          string folder;
          item.GetDisplayName(0x80058000u, out folder); // SIGDN_FILESYSPATH
          return folder;
        } finally {
          Marshal.FinalReleaseComObject(item);
        }
      } finally {
        owner.Close();
        owner.Dispose();
        Marshal.FinalReleaseComObject(dialog);
      }
    }
  }
}
'@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $pickerSource -ReferencedAssemblies @('System.Windows.Forms.dll', 'System.Drawing.dll')

if ($Mode -eq 'workbook') {
  $result = Get-WorkbookPath $WorkbookName
  if ($result) { [Console]::WriteLine($result) }
  exit 0
}
if ($Mode -eq 'pick') {
  $result = Show-WorkspacePicker
  if ($result) { [Console]::WriteLine($result) }
  exit 0
}
# Keep PowerShell and the picker loaded. The bridge sends JSON lines over stdin.
[Console]::WriteLine('{"ready":true}')
while ($true) {
  $line = [Console]::ReadLine()
  if ($null -eq $line) { break }
  try {
    $request = ConvertFrom-Json $line
    $result = if ($request.mode -eq 'pick') { Show-WorkspacePicker } else { Get-WorkbookPath ([string]$request.workbookName) }
    [Console]::WriteLine((ConvertTo-Json -Compress -InputObject @{ id = [string]$request.id; path = $result; error = $null }))
  } catch {
    [Console]::WriteLine((ConvertTo-Json -Compress -InputObject @{ id = [string]$request.id; path = $null; error = $_.Exception.Message }))
  }
}
