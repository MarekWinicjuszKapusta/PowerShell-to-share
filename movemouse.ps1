Add-Type -AssemblyName System.Windows.Forms
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(int flags, int dx, int dy, int cButtons, int info);' -Name U32 -Namespace W;
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 1)
}
Hide-Console
while ($true)
{
  $Pos = [System.Windows.Forms.Cursor]::Position
  $x = $Pos.X + 1
  $y = $Pos.Y + 1
  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
  [W.U32]::mouse_event(6,0,0,0,0);
  Start-Sleep -Seconds 8
}
