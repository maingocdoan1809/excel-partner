/** Fixed script; workbook identity arrives as JSON on stdin, never as code. */
export const WINDOWS_SELECTION_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$request = [Console]::In.ReadLine() | ConvertFrom-Json
$proxies = [Collections.Generic.List[object]]::new()
function Keep($value) {
  if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { $proxies.Add($value) }
  return ,$value
}
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class ExcelSelectionWindow {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr param);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr param);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
  [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
  public static IntPtr[] Children(IntPtr parent) {
    var handles = new List<IntPtr>();
    EnumChildWindows(parent, (hwnd, param) => {
      var name = new StringBuilder(256);
      GetClassName(hwnd, name, 256);
      if (name.ToString() == "EXCEL7") handles.Add(hwnd);
      return true;
    }, IntPtr.Zero);
    return handles.ToArray();
  }
  public static object Native(IntPtr hwnd) {
    var iid = new Guid("00020400-0000-0000-C000-000000000046");
    object value;
    Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out value));
    return value;
  }
}
'@
try {
  # Bind to the actual workbook window, not a stale Running Object Table entry.
  $matches = @()
  foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
    if ($process.MainWindowHandle -eq 0) { continue }
    $title = [string]$process.MainWindowTitle
    $shortName = [IO.Path]::GetFileNameWithoutExtension([string]$request.workbookName)
    if ($title.IndexOf($shortName, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
    foreach ($handle in [ExcelSelectionWindow]::Children($process.MainWindowHandle)) {
      try {
        $candidate = Keep ([ExcelSelectionWindow]::Native($handle))
        $candidateBook = Keep ($candidate.Parent)
        $candidateSheet = Keep ($candidate.ActiveSheet)
        if ($candidateBook.Name -ne $request.workbookName -or $candidateSheet.Name -ne $request.sheet) { continue }
        if ($request.workbookPath -and ([string]$candidateBook.FullName).Replace('/', '\') -ne $request.workbookPath) { continue }
        $matches += $candidate
      } catch { }
    }
  }
  if ($matches.Count -ne 1) { throw 'No unique Excel window matches the task pane workbook and worksheet.' }
  $window = $matches[0]
  $book = Keep ($window.Parent)
  $sheet = Keep ($window.ActiveSheet)
  $selection = Keep ($window.Selection)
  $range = $null
  try { $range = Keep ($selection.ShapeRange) } catch { }
  if ($null -eq $range) {
    @{ status = 'not_shapes' } | ConvertTo-Json -Compress
  } else {
    if ($range.Count -gt 1000) { throw 'Selection exceeds 1000 shapes.' }
    $shapes = @(for ($i = 1; $i -le $range.Count; $i++) {
      $shape = Keep ($range.Item($i))
      @{ name = [string]$shape.Name; nativeId = [string]$shape.ID; nativeType = [int]$shape.Type; connector = ($shape.Connector -eq -1) }
    })
    @{ status = 'ok'; workbookName = [string]$book.Name; sheet = [string]$sheet.Name; shapes = $shapes } | ConvertTo-Json -Depth 4 -Compress
  }
} catch {
  @{ status = 'unavailable'; reason = $_.Exception.Message } | ConvertTo-Json -Compress
} finally {
  for ($i = $proxies.Count - 1; $i -ge 0; $i--) {
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($proxies[$i]) } catch { }
  }
}
`;
