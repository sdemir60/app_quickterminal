Add-Type -Namespace Util -Name WinApi  -MemberDefinition @'
  // Find a window by class name and optionally also title.
  // The TOPMOST matching window (in terms of Z-order) is returned.
  // IMPORTANT: To not search for a title, pass [NullString]::Value, not $null, to lpWindowName
  [DllImport("user32.dll")]
  public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
'@

# Get the topmost File Explorer window, by its class name.
$hwndTopMostFileExplorer = [Util.WinApi]::FindWindow(
  "CabinetWClass",     # the window class of interest
  [NullString]::Value  # no window title to search for
)

if (-not $hwndTopMostFileExplorer) {
  Write-Warning "There is no open File Explorer window."
  # Alternatively, use a *default* directory in this case.
  return
}

# Using a Shell.Application COM object, locate the window by its hWnd and query its location.
$fileExplorerWin = (New-Object -ComObject Shell.Application).Windows() |
                     Where-Object hwnd -eq $hwndTopMostFileExplorer

# This should normally not happen.
if (-not $fileExplorerWin) {
  Write-Warning "The topmost File Explorer window, $hwndTopMostFileExplorer, must have just closed."
  return
}

# Determine the window's active directory (folder) path.
$fileExplorerDir = $fileExplorerWin.Document.Folder.Self.Path

cd $fileExplorerDir
