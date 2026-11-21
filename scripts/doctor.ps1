param(
  [switch]$Full    # also build a one-page presentation from the kit and run it as a hidden slide show (about a minute)
)
# Tells whether this PC can use the skill, and how much of it. Reads only: nothing is installed or changed.
# Every line starts with OK, NOTE (works, with a limit) or PROBLEM (a part will not work), followed by what it means.
# Exit codes: 0 the skill can be used here; 3 it cannot; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\edge.ps1')
. (Join-Path $PSScriptRoot 'lib\fonts.ps1')
Initialize-Console
$script:bad = 0; $script:limits = 0
function Ok([string]$s) { Say "OK       $s" }
function Note([string]$s) { $script:limits++; Say "NOTE     $s" }
function Problem([string]$s) { $script:bad++; Say "PROBLEM  $s" }
function RegValue([string]$Key, [string]$Name) { try { return (Get-ItemProperty -LiteralPath $Key -ErrorAction Stop).$Name } catch { return $null } }
$exit = 0
try {
  Say "ppt-html check of this PC  (skill folder: $($script:HdSkill))"
  Say ''
  # ---- Windows / PowerShell
  $os = [Environment]::OSVersion.Version
  if ($os.Major -ge 10) { Ok "Windows $(if ($os.Build -ge 22000) { '11' } else { '10' }) (build $($os.Build))" } else { Problem "Windows $os : Windows 10 or 11 is needed." }
  $mode = [string]$ExecutionContext.SessionState.LanguageMode
  if ($mode -eq 'FullLanguage') { Ok "Windows PowerShell $($PSVersionTable.PSVersion) (full language mode)" }
  else { Problem "PowerShell runs in '$mode' mode (a security policy of this PC): the scripts of the skill cannot run here." }
  $pol = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -in 'MachinePolicy', 'UserPolicy' -and $_.ExecutionPolicy -in 'AllSigned', 'Restricted' })
  if ($pol.Count -gt 0) { Problem "A group policy forces the script execution policy '$($pol[0].ExecutionPolicy)': unsigned scripts such as these are refused." }
  try { Initialize-HtmlPrep; Ok 'The helper code compiles (C# through Add-Type).' } catch { Problem "The helper code could not be compiled: $($_.Exception.Message)" }
  try {
    $root = Get-HdRoot; $probe = Join-Path $root ('probe_' + [Guid]::NewGuid().ToString('N').Substring(0, 6) + '.txt')
    [IO.File]::WriteAllText($probe, 'x'); Remove-Item -LiteralPath $probe -Force
    Ok "Work folder is writable: $root"
  } catch { Problem "The work folder under %LOCALAPPDATA%\ppt-html cannot be written: $($_.Exception.Message)" }

  # ---- PowerPoint
  $ppt = Find-PowerPointExe
  $major = 16
  if (-not $ppt) { Problem 'PowerPoint (the Windows desktop app) was not found: presentations cannot be built or converted on this PC. (Checking pages with check-html.ps1 still works.)' }
  else {
    $fv = [Diagnostics.FileVersionInfo]::GetVersionInfo($ppt); $major = $fv.FileMajorPart
    $bits = '?'
    try {
      $fs = [IO.File]::Open($ppt, 'Open', 'Read', 'ReadWrite'); $br = New-Object IO.BinaryReader($fs)
      $fs.Position = 0x3C; $pe = $br.ReadInt32(); $fs.Position = $pe + 4; $machine = $br.ReadUInt16(); $br.Dispose(); $fs.Dispose()
      $bits = if ($machine -eq 0x8664 -or $machine -eq 0xAA64) { '64-bit' } else { '32-bit' }
    } catch {}
    if ($major -ge 15) { Ok "PowerPoint $($fv.FileVersion) $bits  ($ppt)" } else { Problem "PowerPoint $($fv.FileVersion) is too old (2013 or newer is needed)." }
    if ($bits -eq '32-bit') { Note '32-bit PowerPoint: supported by the host code, but never measured by the authors of this skill.' }
    # macro settings (what decides whether the live HTML runs when a presentation is opened on THIS PC)
    $ver = "$major.0"
    $polKey = "HKCU:\Software\Policies\Microsoft\Office\$ver\PowerPoint\Security"; $usrKey = "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security"
    $pv = RegValue $polKey 'VBAWarnings'; $uv = RegValue $usrKey 'VBAWarnings'
    $eff = if ($pv -ne $null) { [int]$pv } elseif ($uv -ne $null) { [int]$uv } else { 2 }
    $who = if ($pv -ne $null) { 'set by a policy of the organisation' } else { 'the user setting' }
    switch ($eff) {
      1 { Ok "Macros: enabled without asking ($who)." }
      2 { Ok "Macros: PowerPoint asks once per file (yellow bar, [Enable Content]) - the normal setting ($who)." }
      3 { Note "Macros: only digitally signed macros run ($who). The live HTML of presentations made here will NOT run on this PC; the still picture shows instead. The web version works." }
      4 { Note "Macros: all macros are disabled ($who). Presentations can be built here, but the live HTML will not run on this PC (still picture instead, and the slide show test of the build reports NOT LIVE). The web version works." }
      default { Note "Macros: unknown setting value $eff." }
    }
    if ((RegValue $polKey 'blockcontentexecutionfrominternet') -eq 1) { Note 'Policy: macros of files that came from the internet / mail are blocked. Received .pptm files need [Unblock] in their file properties first (the read-me of the share bundle explains it).' }
  }

  # ---- the engines of the slide show: the modern one (Edge WebView2) and the classic one (Internet Explorer 11)
  $wvVer = ''
  foreach ($k in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}') {
    $v = RegValue $k 'pv'; if ($v -and "$v" -ne '0.0.0.0') { $wvVer = [string]$v; break }
  }
  $loader = $false; if ($ppt) { $loader = Test-Path -LiteralPath (Join-Path (Split-Path -Parent $ppt) 'WebView2Loader.dll') }
  $ie = (Test-Path -LiteralPath (Join-Path $env:WINDIR 'System32\ieframe.dll')) -and (Test-Path -LiteralPath (Join-Path $env:WINDIR 'System32\mshtml.dll'))
  $modern = ($wvVer -and $loader)
  if ($modern) { Ok "The modern page engine is here (Edge WebView2 $wvVer, and Office's loader next to PowerPoint): in the slide show pages run as they were written for a current browser." }
  elseif ($wvVer) { Note "Edge WebView2 $wvVer is installed, but this Office has no WebView2Loader.dll next to POWERPNT.EXE: slide shows on this PC use the classic engine (IE11) with the converted copy of each page." }
  else { Note 'The Edge WebView2 runtime was not found: slide shows on this PC use the classic engine (IE11) with the converted copy of each page. (WebView2 is part of Windows 11 and comes with current Edge and Office.)' }
  if ($ie) { Ok "The classic page engine (the Internet Explorer 11 components of Windows) is here$(if ($modern) { ': the fallback' })." }
  elseif ($modern) { Note 'The Internet Explorer components of Windows (mshtml.dll, ieframe.dll) are missing: there is no fallback engine on this PC (the modern one is used anyway).' }
  else { Problem 'Neither page engine is available (no WebView2 and no Internet Explorer components): pages cannot run inside the slide show on this PC.' }

  # ---- Edge as build tool
  $work = New-WorkDir 'doctor'
  $edge = Start-EdgeTool $work
  if ($edge) { Ok "Edge works as the build tool ($($edge.Browser)): pictures of the pages as the modern engine draws them, font embedding, the converted copies for the classic engine and the check of the web version are available."; Stop-EdgeTool }
  else { Note "No build browser ($(Get-EdgeWhyNot)). Presentations can then only be built for the classic engine: pages written for IE11 (the design kit and its example pages are) build as before; modern pages are not converted, fonts are not embedded, the web version is not opened for a check." }

  # ---- fonts
  $idx = Get-FontIndex
  foreach ($fam in 'Noto Sans KR', 'Noto Serif KR') {
    $k = $fam.ToLowerInvariant()
    if (-not $idx.ContainsKey($k)) { Note "Font '$fam' is neither installed nor in the font pack: the kit falls back to Malgun Gothic / Batang." ; continue }
    $sys = @($idx[$k] | Where-Object { -not $_.Bundled })
    if ($sys.Count -gt 0) { Ok "Font '$fam' is installed (embedded into pages from here)." } else { Ok "Font '$fam' comes from the skill's font pack (2,350 everyday Hangul syllables; rarer syllables fall back to a system font)." }
  }

  # ---- can PowerShell read the skill's own files (security software sometimes refuses single files)
  $unread = @()
  foreach ($f in Get-ChildItem -LiteralPath (Join-Path $script:HdSkill 'assets') -Recurse -File) {
    try { $rs = [IO.File]::OpenRead($f.FullName); try { [void]$rs.ReadByte() } finally { $rs.Dispose() } } catch { $unread += $f.FullName.Substring($script:HdSkill.Length + 1) }
  }
  if ($unread.Count -eq 0) { Ok 'Every file of the kit, the templates and the build tools can be read.' }
  else { Note ("This PC does not let PowerShell read $($unread.Count) file(s) of the skill (security software is the usual cause): " + (($unread | Select-Object -First 8) -join ', ') + '. A theme among them cannot be used from the skill folder until that changes; the scripts leave such files out and say so.') }

  # ---- the real thing
  if ($Full) {
    Say ''
    Say 'Full test: building one kit page into a presentation and running it as a hidden slide show ...'
    if (-not $ppt) { Problem 'Full test skipped: no PowerPoint.' }
    else {
      $outFile = Join-Path $work 'doctor.pptm'
      $log = Join-Path $work 'doctor-build.txt'
      $p = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + (Join-Path $PSScriptRoot 'html2ppt.ps1') + '"'), '-Html', ('"' + (Join-Path $script:HdSkill 'assets\templates\kpi.html') + '"'), '-Out', ('"' + $outFile + '"'), '-Web') -WindowStyle Hidden -PassThru -Wait -RedirectStandardOutput $log
      $tail = @(); try { $tail = @(Get-Content -LiteralPath $log -Encoding UTF8 | Where-Object { $_ -match 'slide 1:|fallback|checked in Edge|DONE|FAILED|WARNING|ERROR' }) } catch {}
      if ($p.ExitCode -eq 0) { Ok 'Full test passed: the page was live in a real (hidden) slide show (the line below names the engine) and the web version ran in Edge.' }
      elseif ($p.ExitCode -eq 4) { Note 'Full test: the presentation was built, but the live page was not seen in the slide show on this PC (macro setting or security software). Presentations built here still work on other PCs.' }
      else { Problem "Full test failed (exit code $($p.ExitCode)). Log: $log" }
      foreach ($t in $tail) { Say "         $t" }
    }
  }

  Say ''
  if ($script:bad -gt 0) { $exit = 3; Say "RESULT: $($script:bad) problem(s). The parts named above will not work on this PC." }
  elseif ($script:limits -gt 0) { Say "RESULT: usable, with $($script:limits) limit(s) (NOTE lines)." }
  else { Say 'RESULT: everything the skill needs is here.' }
  if (-not $Full) { Say '        (Add -Full to build and run a real test presentation.)' }
} catch {
  $exit = 1
  Say "FAILED: $($_.Exception.Message)"
  Say $_.ScriptStackTrace
} finally { try { Stop-EdgeTool } catch {} }
exit $exit
