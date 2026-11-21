# Dot-source (entry scripts). Work folders, the hidden PowerPoint instance, child scripts, delivery of results.
# Plain ASCII on purpose (see hd_core.ps1).
. (Join-Path $PSScriptRoot 'hd_core.ps1')

$script:Utf8NoBom = New-Object Text.UTF8Encoding($false)

function Initialize-Console {
  try { [Console]::OutputEncoding = $script:Utf8NoBom } catch {}
}
function Say([string]$Text) { [Console]::Out.WriteLine($Text) }

function New-WorkDir([string]$Kind) {
  $root = Join-Path (Get-HdRoot) 'work'
  if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Force $root | Out-Null }
  Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-3) } | ForEach-Object {
    try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop } catch {}
  }
  $d = Join-Path $root ("{0}-{1}-{2}" -f $Kind, (Get-Date -Format 'yyyyMMdd-HHmmss'), ([Guid]::NewGuid().ToString('N').Substring(0, 4)))
  New-Item -ItemType Directory -Force $d | Out-Null
  return $d
}

function Invoke-Retry([scriptblock]$Action, [int]$Times = 5, [int]$DelayMs = 500) {
  for ($i = 1; $i -le $Times; $i++) {
    try { return (& $Action) } catch { if ($i -eq $Times) { throw }; Start-Sleep -Milliseconds $DelayMs }
  }
}

function Read-Asset([string]$Name) {
  return [IO.File]::ReadAllText((Join-Path $script:HdSkill ('assets\' + $Name)), [Text.Encoding]::UTF8).Trim()
}

function Initialize-HtmlPrep {
  if (-not ('PptHtml.HtmlPrep' -as [type])) { Add-Type -AssemblyName System.Drawing; Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\HtmlPrep.cs'))) -ReferencedAssemblies System.Drawing }
}

# Writes the VBA project of the inline-HTML host (the text sources in scripts\vba) to a file. The project
# holds source only; Office compiles it when the presentation is opened. (Access to the VBA object model is
# normally blocked, so the file format is written directly.)
function New-VbaProject([string]$OutFile) {
  if (-not ('OleKit.VbaProj' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\VbaKit.cs'))) }
  $src = Join-Path $script:HdScripts 'vba'
  $mods = @('modInline.bas', 'modWv2.bas', 'clsAppEvents.cls', 'clsDomBridge.cls', 'frmInline.frm')
  foreach ($m in $mods) {
    $t = [IO.File]::ReadAllText((Join-Path $src $m))
    foreach ($ch in $t.ToCharArray()) { if ([int]$ch -gt 127) { throw "VBA source $m contains a non-ASCII character" } }
    foreach ($line in ($t -split "`n")) { if ($line.TrimEnd("`r").Length -gt 1000) { throw "VBA source $m has a line longer than 1000 characters" } }
  }
  $pf = $env:ProgramFiles
  $refs = @(
    ('VBA|*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#' + $pf + '\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications'),
    ('PowerPoint|*\G{91493440-5A91-11CF-8700-00AA0060263B}#2.c#0#' + $pf + '\Microsoft Office\root\Office16\MSPPT.OLB#Microsoft PowerPoint 16.0 Object Library'),
    'stdole|*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation',
    ('Office|*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#' + $pf + '\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library'),
    'MSForms|*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Windows\System32\FM20.DLL#Microsoft Forms 2.0 Object Library'
  )
  return [OleKit.VbaProj]::Build($OutFile, $src, [string[]]$mods, 'VBAProject', [string[]]$refs, 949, [uint32]0x0412, [uint32]3, 480, 176, 'InlineHtmlHost')
}

function Find-PowerPointExe {
  foreach ($k in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe') {
    try {
      $v = (Get-ItemProperty -LiteralPath $k -ErrorAction Stop).'(default)'
      if ($v) { $v = $v.Trim('"'); if (Test-Path -LiteralPath $v) { return $v } }
    } catch {}
  }
  $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
  foreach ($r in $roots) {
    foreach ($sub in 'Microsoft Office\root\Office16', 'Microsoft Office\Office16', 'Microsoft Office\root\Office15', 'Microsoft Office\Office15') {
      $p = Join-Path $r (Join-Path $sub 'POWERPNT.EXE')
      if (Test-Path -LiteralPath $p) { return $p }
    }
  }
  return $null
}

# Starts the hidden PowerPoint instance (or reuses a running one). Returns 'started' or 'reused'; throws when it cannot.
function Start-Hd([int]$WaitSec = 60) {
  if (Get-HdPid) { try { [void](Get-HdApp); return 'reused' } catch {} }
  $exe = Find-PowerPointExe
  if (-not $exe) { throw 'PowerPoint (the Windows desktop app) was not found on this PC.' }
  $dir = Join-Path (Get-HdRoot) 'hd'
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
  # The modern page engine (WebView2) of the hidden instance gets a browser profile of its own: it must never
  # share one with a PowerPoint the user has open (the host code in the presentation reads this variable).
  $env:PPTHTML_WV2_DATA = Join-Path $dir 'wvdata'
  # the host code writes its log lines to this file as they happen (what it did last when something goes wrong)
  $env:PPTHTML_LOG = Join-Path $dir 'host.log'
  try { if (Test-Path -LiteralPath $env:PPTHTML_LOG) { Remove-Item -LiteralPath $env:PPTHTML_LOG -Force } } catch {}
  # /B = start with a new blank presentation: it gives the instance a document window to be reached through
  $newPid = [DeskKit.Desk]::Start($script:HdDesk, ('"' + $exe + '" /B'), $dir)
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $alive = $true
  while ($sw.Elapsed.TotalSeconds -lt $WaitSec) {
    Start-Sleep -Milliseconds 600
    try {
      $app = Get-HdApp
      if ($app.Presentations.Count -ge 1) {
        # The blank presentation is the anchor of the instance. It is deliberately left "changed": PowerPoint
        # closes an untouched blank presentation by itself as soon as another file is opened in a window.
        $anchor = $app.Presentations.Item(1)
        [void]$anchor.Tags.Add('PptHtmlAnchor', '1')
        try { [void]$anchor.Slides.Add($anchor.Slides.Count + 1, 12) } catch {}
        return 'started'
      }
    } catch {}
    $alive = [bool](Get-Process -Id $newPid -ErrorAction SilentlyContinue)
    if (-not $alive) { break }
    foreach ($w in [DeskKit.Desk]::Windows($script:HdDesk, $newPid)) {
      if ($w.Visible -and ($w.Class -eq '#32770' -or $w.Class -eq 'NUIDialog')) { [void][DeskKit.Desk]::Close($script:HdDesk, $w.Hwnd) }
    }
  }
  if ($alive) {
    try { Stop-Process -Id $newPid -Force -ErrorAction SilentlyContinue } catch {}
    throw 'A separate, hidden PowerPoint could not be started: it did not become ready in time.'
  }
  throw 'A separate, hidden PowerPoint could not be started: the new process ended at once. If PowerPoint is open, it may have taken the request itself (an empty presentation may have appeared there). Close PowerPoint and run again.'
}

# Ends the hidden instance. Only ever touches the process that owns the windows on the hidden desktop.
function Close-HdApp {
  try {
    $app = Get-HdApp
    try { for ($i = $app.SlideShowWindows.Count; $i -ge 1; $i--) { $app.SlideShowWindows.Item($i).View.Exit() } } catch {}
    for ($i = $app.Presentations.Count; $i -ge 1; $i--) {
      try { $p = $app.Presentations.Item($i); $p.Saved = -1; $p.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p) } catch {}
    }
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
  } catch {}
}
function Stop-Hd {
  $procId = Get-HdPid
  if (-not $procId) { return 'not running' }
  Close-HdApp
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  $sw = [Diagnostics.Stopwatch]::StartNew()
  while ($sw.Elapsed.TotalSeconds -lt 6 -and (Get-Process -Id $procId -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
  if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { try { Stop-Process -Id $procId -Force -ErrorAction Stop } catch {}; return 'ended (forced)' }
  return 'ended'
}

# Makes sure the hidden instance is there and answers; restarts it when it does not.
function Assert-Hd {
  try { [void](Get-HdApp); return } catch {}
  try { [void](Stop-Hd) } catch {}
  [void](Start-Hd)
}

# Runs lib\<Script> in its own PowerShell process with the given arguments (hashtable -> JSON file).
# While it runs, dialogs that pop up in the hidden PowerPoint (VBA errors, message boxes) are recorded and
# closed so that automation never hangs. Returns an object: Ok, Result (the child's JSON), Log, Dialogs, TimedOut.
function Invoke-HdChild {
  param([Parameter(Mandatory)][string]$Script, [hashtable]$Params = @{}, [Parameter(Mandatory)][string]$WorkDir, [int]$TimeoutSec = 300, [switch]$NoHd)
  $runDir = Join-Path $WorkDir 'run'
  if (-not (Test-Path -LiteralPath $runDir)) { New-Item -ItemType Directory -Force $runDir | Out-Null }
  $id = [IO.Path]::GetFileNameWithoutExtension($Script) + '_' + [Guid]::NewGuid().ToString('N').Substring(0, 6)
  $argsFile = Join-Path $runDir "$id.args.json"
  $log = Join-Path $runDir "$id.log.txt"
  $result = Join-Path $runDir "$id.result.json"
  $stdout = Join-Path $runDir "$id.out.txt"
  $stderr = Join-Path $runDir "$id.err.txt"
  $Params['log'] = $log
  $Params['result'] = $result
  Write-JsonFile $argsFile $Params
  if (-not $NoHd) { Assert-Hd }
  $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', ('"' + (Join-Path $script:HdLib $Script) + '"'), '-ArgsFile', ('"' + $argsFile + '"'))
  $child = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $argList -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $dialogs = New-Object Collections.Generic.List[string]
  $timedOut = $false
  $hdPid = 0
  if (-not $NoHd) { $hdPid = Get-HdPid }
  while (-not $child.HasExited) {
    Start-Sleep -Milliseconds 300
    if ($hdPid) {
      foreach ($w in [DeskKit.Desk]::Windows($script:HdDesk, $hdPid)) {
        if ($w.Visible -and ($w.Class -eq '#32770' -or $w.Class -eq 'NUIDialog')) {
          $dialogs.Add("[$($w.Title)] $($w.ChildText)")
          try { [void](Save-HdShot $w.Hwnd (Join-Path $runDir ("$id.dialog$($dialogs.Count).png")) 0) } catch {}
          [void][DeskKit.Desk]::Close($script:HdDesk, $w.Hwnd)
          Start-Sleep -Milliseconds 400
          if ($w.Title -match 'Visual Basic') {
            Start-Sleep -Milliseconds 600
            $vbe = Get-HdWindows | Where-Object { $_.Class -eq 'wndclass_desked_gsk' } | Select-Object -First 1
            if ($vbe) { try { [void](Save-HdShot $vbe.Hwnd (Join-Path $runDir ("$id.vbe$($dialogs.Count).png")) 0) } catch {} }
            Reset-HdVbe
          }
        }
      }
    }
    if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { $timedOut = $true; try { $child.Kill() } catch {}; break }
  }
  $logText = ''
  if (Test-Path -LiteralPath $log) { $logText = [IO.File]::ReadAllText($log, [Text.Encoding]::UTF8) }
  foreach ($f in $stderr, $stdout) {
    if ((Test-Path -LiteralPath $f) -and (Get-Item -LiteralPath $f).Length -gt 0) { $logText += "`r`n" + [IO.File]::ReadAllText($f) }
  }
  $res = $null
  if (Test-Path -LiteralPath $result) { try { $res = Read-JsonFile $result } catch { $logText += "`r`nresult file unreadable: $($_.Exception.Message)" } }
  return [pscustomobject]@{
    Ok = ($res -ne $null -and -not $timedOut -and -not $res.error)
    Result = $res
    Log = $logText.Trim()
    Dialogs = @($dialogs)
    TimedOut = $timedOut
    Seconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
  }
}

# Describes why a child failed, for the console.
function Format-ChildFailure($Run, [string]$What) {
  $lines = @("$What failed.")
  if ($Run.TimedOut) { $lines += 'It did not finish in time.' }
  if ($Run.Result -and $Run.Result.error) { $lines += "Error: $($Run.Result.error)" }
  foreach ($d in $Run.Dialogs) { $lines += "PowerPoint showed a dialog: $d" }
  if ($Run.Log) { $lines += 'Log:'; $lines += $Run.Log }
  return ($lines -join "`r`n")
}

# Copies a finished file to where the user wants it. If that name is in use (the file is open), a numbered
# name is tried. Returns the path written, or $null when the folder cannot be written at all.
function Copy-Deliver([string]$Source, [string]$Dest) {
  $dir = Split-Path -Parent $Dest
  try { if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null } } catch {}
  $name = [IO.Path]::GetFileNameWithoutExtension($Dest)
  $ext = [IO.Path]::GetExtension($Dest)
  for ($i = 1; $i -le 4; $i++) {
    $target = if ($i -eq 1) { $Dest } else { Join-Path $dir ("{0} ({1}){2}" -f $name, $i, $ext) }
    try { [IO.File]::Copy($Source, $target, $true); return $target } catch {}
  }
  return $null
}

function Resolve-UserPath([string]$Path, [string]$BaseDir) {
  if (-not $Path) { return $null }
  $p = $Path.Trim().Trim('"')
  if ([IO.Path]::IsPathRooted($p)) { return [IO.Path]::GetFullPath($p) }
  return [IO.Path]::GetFullPath((Join-Path $BaseDir $p))
}

# Adds the macro project (and the ribbon part whose onLoad callback starts the code when the file is opened)
# to a macro-enabled presentation package. Works on a copy; returns notes about what was done.
function Add-VbaToPackage([string]$Base, [string]$Vba, [string]$Out) {
  Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
  $notes = @()
  Invoke-Retry { [IO.File]::Copy($Base, $Out, $true) } | Out-Null
  $fs = [IO.File]::Open($Out, 'Open', 'ReadWrite')
  $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Update)
  $readEntry = {
    param($name)
    $e = $zip.GetEntry($name); if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Dispose(); return $t
  }
  $writeEntry = {
    param($name, [string]$text)
    $e = $zip.GetEntry($name); if ($e) { $e.Delete() }
    $e = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
    $s = $e.Open(); $b = $script:Utf8NoBom.GetBytes($text); $s.Write($b, 0, $b.Length); $s.Dispose()
  }
  try {
    $ct = & $readEntry '[Content_Types].xml'
    if ($ct -notmatch 'presentation\.macroEnabled\.main\+xml') { throw 'the staged file is not a macro-enabled presentation' }
    if ($ct -notmatch 'vnd\.ms-office\.vbaProject') {
      $ct = $ct -replace '</Types>', '<Override PartName="/ppt/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'
      & $writeEntry '[Content_Types].xml' $ct
    }
    $rels = & $readEntry 'ppt/_rels/presentation.xml.rels'
    if ($rels -notmatch 'relationships/vbaProject') {
      $rels = $rels -replace '</Relationships>', '<Relationship Id="rIdVbaProject1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'
      & $writeEntry 'ppt/_rels/presentation.xml.rels' $rels
    }
    $target = 'ppt/vbaProject.bin'
    if ($rels -match 'relationships/vbaProject"\s+Target="([^"]+)"') { $target = 'ppt/' + $Matches[1].TrimStart('/') }
    $e = $zip.GetEntry($target); if ($e) { $e.Delete() }
    $e = $zip.CreateEntry($target, [IO.Compression.CompressionLevel]::Optimal)
    $s = $e.Open(); $b = [IO.File]::ReadAllBytes($Vba); $s.Write($b, 0, $b.Length); $s.Dispose()
    $notes += "macro project added ($($b.Length) bytes)"
    # ribbon part: <customUI onLoad=...> runs InlineHtml_OnLoad when the file is opened with macros enabled
    $cui = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n" + '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="InlineHtml_OnLoad"><ribbon/></customUI>'
    $root = & $readEntry '_rels/.rels'
    $existing = [regex]::Matches($root, 'relationships/ui/extensibility"\s+Target="([^"]+)"')
    $foreign = $false
    foreach ($m in $existing) {
      $t = & $readEntry ($m.Groups[1].Value.TrimStart('/'))
      if ($t -and $t -notmatch 'InlineHtml_OnLoad') { $foreign = $true }
    }
    if ($foreign) {
      $notes += 'WARNING: the presentation has its own ribbon customisation; the start-on-open hook was not added (the page starts at the first slide change or mouse move instead)'
    } else {
      & $writeEntry 'customUI/customUI14.xml' $cui
      if ($root -notmatch 'ui/extensibility') {
        $root = $root -replace '</Relationships>', '<Relationship Id="rIdPptHtmlUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
        & $writeEntry '_rels/.rels' $root
      }
      $notes += 'start-on-open hook added'
    }
  } finally { $zip.Dispose(); $fs.Dispose() }
  return $notes
}

# Reads a text file the user points at. Security software on some PCs refuses script hosts access to files in
# Desktop / Documents: say so plainly instead of failing with a bare "access denied".
function Read-UserBytes([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
  try { return (Invoke-Retry { [IO.File]::ReadAllBytes($Path) } 3 300) }
  catch { throw "ACCESS-DENIED: this PC does not let PowerShell read '$Path'. Copy the file (and the files it refers to) into a folder under $(Get-HdRoot)\in with your own file tools and pass that copy instead." }
}
