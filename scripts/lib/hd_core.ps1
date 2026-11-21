# Dot-source. Helpers for the separate PowerPoint instance that this skill drives.
#
# The instance runs on its own hidden Windows desktop, so nothing appears on screen and a PowerPoint window the
# user has open is never touched. All files of this skill are plain ASCII on purpose: Windows PowerShell 5.1
# reads script files without a byte order mark in the PC's legacy code page.
$script:HdDesk = 'PptHtmlDesk'
$script:HdLib = $PSScriptRoot
$script:HdScripts = Split-Path -Parent $script:HdLib
$script:HdSkill = Split-Path -Parent $script:HdScripts
Add-Type -AssemblyName System.Drawing
if (-not ('DeskKit.Desk' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\DeskKit.cs'))) -ReferencedAssemblies System.Drawing }

function Get-HdRoot {
  $r = Join-Path $env:LOCALAPPDATA 'ppt-html'
  if (-not (Test-Path -LiteralPath $r)) { New-Item -ItemType Directory -Force $r | Out-Null }
  return (Get-Item -LiteralPath $r).FullName
}
function Get-HdFrames { @([DeskKit.Desk]::Windows($script:HdDesk, 0) | Where-Object { $_.Class -eq 'PPTFrameClass' }) }
function Test-HdDesk { return [DeskKit.Desk]::Exists($script:HdDesk) }
function Get-HdPid {
  if (-not (Test-HdDesk)) { return 0 }
  $f = Get-HdFrames | Select-Object -First 1
  if ($f) { return $f.Pid } else { return 0 }
}
# Application object of the hidden instance (through the document window, so it can never be another instance).
function Get-HdApp {
  if (Test-HdDesk) {
    foreach ($f in (Get-HdFrames)) {
      $mdi = [DeskKit.Desk]::FindDescendant($script:HdDesk, $f.Hwnd, 'mdiClass', 0)
      if ($mdi -ne [IntPtr]::Zero) { try { return ([DeskKit.Desk]::NativeOm($script:HdDesk, $mdi)).Application } catch {} }
    }
  }
  throw 'the hidden PowerPoint instance is not reachable'
}
function Get-HdWindows([switch]$All) {
  $procId = Get-HdPid
  if (-not $procId) { return @() }
  [DeskKit.Desk]::Windows($script:HdDesk, $procId) | Where-Object { $All -or $_.Visible }
}
function Save-HdShot([IntPtr]$Hwnd, [string]$Path, [int]$Flags = 2) {
  return [DeskKit.Desk]::Capture($script:HdDesk, $Hwnd, $Path, $Flags)
}
# Runs a macro of a presentation open in the hidden instance (works without access to the VBA project).
function Invoke-HdMacro($App, $Pres, [string]$Name, [object[]]$MacroArgs = @()) {
  $full = "$($Pres.Name)!$Name"
  $callArgs = New-Object Collections.Generic.List[object]
  $callArgs.Add($full)
  foreach ($a in $MacroArgs) { $callArgs.Add($a) }
  try { return [string]($App.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $App, $callArgs.ToArray())) }
  catch { $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }; return "RUN $Name FAILED: $($e.Message)" }
}
# Leaves break mode of the VBA editor of the hidden instance (after a compile or run-time error) and closes it.
function Reset-HdVbe {
  foreach ($vbe in @(Get-HdWindows | Where-Object { $_.Class -eq 'wndclass_desked_gsk' })) {
    $bar = $null
    foreach ($line in [DeskKit.Desk]::Tree($script:HdDesk, $vbe.Hwnd)) {
      if ($line -match '^hwnd=(\d+) class=MsoCommandBar title=\[[^\]]*\] rect=\S+ 996x38') { $bar = [IntPtr][int64]$Matches[1]; break }
    }
    if (-not $bar) { $bar = [DeskKit.Desk]::FindDescendant($script:HdDesk, $vbe.Hwnd, 'MsoCommandBar', 0) }
    [void][DeskKit.Desk]::PostClick($script:HdDesk, $bar, 438, 18)
    Start-Sleep -Milliseconds 700
    [void][DeskKit.Desk]::Close($script:HdDesk, $vbe.Hwnd)
    Start-Sleep -Milliseconds 300
  }
}
# Closes every presentation of the hidden instance except its anchor (the blank presentation it was started
# with; it keeps a document window alive, which is how the instance is reached).
function Close-HdExtra($App) {
  for ($i = $App.Presentations.Count; $i -ge 1; $i--) {
    $p = $App.Presentations.Item($i)
    $anchor = $false
    try { $anchor = ($p.Tags.Item('PptHtmlAnchor') -eq '1') } catch {}
    if (-not $anchor) { try { $p.Saved = -1; $p.Close() } catch {} }
  }
}
function ConvertTo-OleColor([string]$Hex) {
  $h = $Hex.TrimStart('#')
  return [int]([Convert]::ToInt32($h.Substring(0, 2), 16) + 256 * [Convert]::ToInt32($h.Substring(2, 2), 16) + 65536 * [Convert]::ToInt32($h.Substring(4, 2), 16))
}

# ---- conventions of the child scripts (lib\*.ps1 run through Invoke-HdChild) -------------------------------
# A child gets one parameter, -ArgsFile: a UTF-8 JSON file with its arguments plus "log" and "result" paths.
function Read-JsonFile([string]$Path) { return ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json) }
function Write-JsonFile([string]$Path, $Object) {
  $json = $Object | ConvertTo-Json -Depth 20
  [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
}
function Write-ChildLog([string]$Path, [string]$Text) {
  [IO.File]::AppendAllText($Path, $Text + "`r`n", (New-Object Text.UTF8Encoding($false)))
}
