param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): runs the finished presentation as a real slide show, visits every slide
# that carries a page and checks that the live page is there: a browser window inside the slide show window,
# script running, no script errors. It measures how smoothly the page moves (frames per second while it
# appears, and while the test actions run) and saves a picture of the slide show per slide.
#   args: file, shotsDir, showType (1 = full screen, 2 = window), maxShotWidth, perf (false = no frame measurement),
#         click = [ actions ], each for one slide, in canvas coordinates of its page:
#           { slide, x, y }                      a click
#           { slide, x, y, hold: 700 }           press and hold (ms), then release
#           { slide, x, y, to: [x2, y2] }        drag (optional hold: ms before moving)
#           { slide, text: "abc" }               characters, as typed (after a click that put the caret somewhere)
#           { slide, key: 13 }                   a key (virtual key code)
#           { slide, x, y, wheel: 600 }          the mouse wheel over that point, in px down (modern engine only)
#           { slide, eval: "js expression" }     evaluated in the page; the value goes into the report
#         any action may carry  wait: ms  (pause after it, default 900)
#         stills = true saves a picture of each visited slide as it looks OUTSIDE the show afterwards (still_after_NN.png)
#         engine = 'classic' makes the show use the classic engine (the fallback) even where the modern one is there;
#         noClassic = [ page numbers ] that have no document for the classic engine (they are not waited for then)
#         only = [ slide numbers ] visits just those slides; keep = true opens the file for writing and saves it
#         after the show (what the pages stored during the test stays in the file: for testing that, nothing else)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
$pres = $null; $app = $null; $oldSec = $null; $w = $null
function ShowWin { Get-HdWindows | Where-Object { $_.Class -eq 'screenClass' -or ($_.Class -eq 'PPTFrameClass' -and $_.Title -match ' - \[') } | Select-Object -First 1 }
# the browser window of the page on screen: the Internet Explorer_Server whose left edge is its host form's
function CurIe($sw) {
  $form = $null; $ie = $null
  foreach ($line in [DeskKit.Desk]::Tree($script:HdDesk, $sw.Hwnd)) {
    if ($line -match '^hwnd=(\d+) class=ThunderDFrame .* rect=(-?\d+),(-?\d+) (\d+)x(\d+) .* vis=(\w+)') { $form = @{ x = [int]$Matches[2]; y = [int]$Matches[3]; w = [int]$Matches[4]; h = [int]$Matches[5]; vis = $Matches[6] } }
    elseif ($form -and $line -match '^hwnd=(\d+) class=Internet Explorer_Server .* rect=(-?\d+),(-?\d+) (\d+)x(\d+)') {
      if ([int]$Matches[2] -eq $form.x -and $form.vis -eq 'True') { $ie = @{ h = [IntPtr][int64]$Matches[1]; x = [int]$Matches[2]; y = [int]$Matches[3]; w = [int]$Matches[4]; hh = [int]$Matches[5] } }
    }
  }
  return $ie
}
function Shrink([string]$path, [int]$maxW) {
  $img = [Drawing.Image]::FromFile($path)
  try {
    if ($img.Width -le $maxW) { return "$($img.Width)x$($img.Height)" }
    $nh = [int]($img.Height * $maxW / $img.Width)
    $bmp = New-Object Drawing.Bitmap($maxW, $nh)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($img, 0, 0, $maxW, $nh); $g.Dispose()
  } finally { $img.Dispose() }
  $bmp.Save($path, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  return "${maxW}x$nh"
}
function PageEval([int]$page, [string]$js) { return [string](Invoke-HdMacro $app $pres 'InlineHtml_Eval' @($page, $js)) }
# "fps|longest frame ms|frames over 34 ms|frames" -> object (or $null when the page has no counter)
function ReadFps([int]$page) {
  $r = PageEval $page 'window.__ihFps ? window.__ihFps(0) : ""'
  if ($r -notmatch '^(\d+)\|(\d+)\|(\d+)\|(\d+)$') { return $null }
  if ([int]$Matches[4] -lt 3) { return $null }
  return [pscustomobject]@{ fps = [int]$Matches[1]; worst = [int]$Matches[2]; slow = [int]$Matches[3]; frames = [int]$Matches[4] }
}
# ---- input for a page of the modern engine: through the browser's developer protocol, in CSS px of the page
function Num([double]$v) { return [Math]::Round($v, 1).ToString([Globalization.CultureInfo]::InvariantCulture) }
function Cdp([int]$page, [string]$method, [string]$json) { return [string](Invoke-HdMacro $app $pres 'InlineHtml_Cdp' @($page, $method, $json)) }
function MMouse([int]$page, [string]$type, [double]$x, [double]$y, [bool]$down) {
  $moving = ($type -eq 'mouseMoved')
  $button = if ($moving -and -not $down) { 'none' } else { 'left' }
  $buttons = if ($type -eq 'mouseReleased' -or ($moving -and -not $down)) { 0 } else { 1 }
  $j = '{"type":"' + $type + '","x":' + (Num $x) + ',"y":' + (Num $y) + ',"button":"' + $button + '","buttons":' + $buttons + ',"clickCount":' + $(if ($moving) { 0 } else { 1 }) + '}'
  [void](Cdp $page 'Input.dispatchMouseEvent' $j)
}
function MClick([int]$page, [double]$x, [double]$y, [int]$holdMs) {
  MMouse $page 'mouseMoved' $x $y $false; Start-Sleep -Milliseconds 40
  MMouse $page 'mousePressed' $x $y $true; Start-Sleep -Milliseconds ([Math]::Max(40, $holdMs))
  MMouse $page 'mouseReleased' $x $y $false
}
function MDrag([int]$page, [double]$x, [double]$y, [double]$x2, [double]$y2, [int]$steps, [int]$stepMs, [int]$holdMs) {
  MMouse $page 'mouseMoved' $x $y $false; Start-Sleep -Milliseconds 40
  MMouse $page 'mousePressed' $x $y $true; Start-Sleep -Milliseconds ([Math]::Max(40, $holdMs))
  for ($q = 1; $q -le $steps; $q++) { MMouse $page 'mouseMoved' ($x + ($x2 - $x) * $q / $steps) ($y + ($y2 - $y) * $q / $steps) $true; Start-Sleep -Milliseconds $stepMs }
  MMouse $page 'mouseReleased' $x2 $y2 $false
}
function MKey([int]$page, [int]$vk) {
  [void](Cdp $page 'Input.dispatchKeyEvent' ('{"type":"rawKeyDown","windowsVirtualKeyCode":' + $vk + ',"nativeVirtualKeyCode":' + $vk + '}'))
  Start-Sleep -Milliseconds 30
  [void](Cdp $page 'Input.dispatchKeyEvent' ('{"type":"keyUp","windowsVirtualKeyCode":' + $vk + ',"nativeVirtualKeyCode":' + $vk + '}'))
}
function MWheel([int]$page, [double]$x, [double]$y, [double]$dy) {
  MMouse $page 'mouseMoved' $x $y $false; Start-Sleep -Milliseconds 30
  [void](Cdp $page 'Input.dispatchMouseEvent' ('{"type":"mouseWheel","x":' + (Num $x) + ',"y":' + (Num $y) + ',"deltaX":0,"deltaY":' + (Num $dy) + '}'))
}
function MText([int]$page, [string]$text) { [void](Cdp $page 'Input.insertText' ('{"text":' + (ConvertTo-Json $text -Compress) + '}')) }
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 1 } catch {}
  $keep = [bool]$A.keep
  $pres = $app.Presentations.Open([string]$A.file, $(if ($keep) { 0 } else { -1 }), 0, -1)
  Start-Sleep -Milliseconds 800
  # which engine the show may use: a tag of the presentation (set here in memory only) asks for the classic one.
  # (No macro is called before the show: a call that early, while PowerPoint is still loading the code, fails
  # and was seen to leave every later call failing as well.)
  try { $pres.Tags.Delete('HtmlEngine') } catch {}
  if ([string]$A.engine -eq 'classic') { [void]$pres.Tags.Add('HtmlEngine', 'classic') }
  $noClassic = @(); if ($A.noClassic) { $noClassic = @($A.noClassic | ForEach-Object { [int]$_ }) }
  $modern = $false; $browser = ''
  $targets = @()
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $tag = ''; try { $tag = [string]$pres.Slides.Item($i).Tags.Item('HtmlPage') } catch {}
    if ($tag) { $targets += [pscustomobject]@{ slide = $i; page = [int]$tag } }
  }
  if ($A.only) { $only = @($A.only | ForEach-Object { [int]$_ }); $targets = @($targets | Where-Object { $only -contains $_.slide }) }
  $showType = 1; if ($A.showType) { $showType = [int]$A.showType }
  $maxW = 1600; if ($A.maxShotWidth) { $maxW = [int]$A.maxShotWidth }
  $perf = $true; if ($A.perf -ne $null -and -not [bool]$A.perf) { $perf = $false }
  $ss = $pres.SlideShowSettings; $ss.ShowType = $showType
  $startClock = [Diagnostics.Stopwatch]::StartNew()
  $w = $ss.Run()
  Start-Sleep -Milliseconds 2500
  $sw = ShowWin
  if (-not $sw) { throw 'the slide show window did not appear' }
  $showDesc = [DeskKit.Desk]::Describe($script:HdDesk, $sw.Hwnd)
  $results = @(); $firstLive = 0.0
  foreach ($t in $targets) {
    try { if ($w.View.Slide.SlideIndex -ne $t.slide) { $w.View.GotoSlide($t.slide) } } catch { $w.View.GotoSlide($t.slide) }
    $ie = $null; $status = ''
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while ($clock.Elapsed.TotalSeconds -lt 20) {
      Start-Sleep -Milliseconds 400
      $status = ((Invoke-HdMacro $app $pres 'InlineHtml_Status') -split "`r`n")[0]
      $modern = ($status -match ' engine=modern ')
      if ($status -match (' cur=' + $t.page + ' ') -and $status -match ' shown=True ') {
        if ($modern) { if ($status -match (' p' + $t.page + '\[alive,hooked')) { $ie = @{ modern = $true; w = 0 }; break } }
        else { $ie = CurIe $sw; if ($ie) { break } }
      }
      # the classic engine has nothing to show for this page: its still picture is all there will be
      if (-not $modern -and $noClassic -contains $t.page -and $clock.Elapsed.TotalSeconds -gt 4) { break }
    }
    if ($modern -and $ie -and -not $browser) { $browser = PageEval $t.page '(navigator.userAgent.match(/Edg\/(\d+)/)||[0,""])[1]' }
    if ($firstLive -eq 0 -and $ie) { $firstLive = [Math]::Round($startClock.Elapsed.TotalSeconds, 1) }
    Start-Sleep -Milliseconds 700
    $pageInfo = ''
    if ($status -match (' p' + $t.page + '\[([^\]]*)\]')) { $pageInfo = $Matches[1] }
    $view = ''; if ($pageInfo -match 'view=(\d+x\d+)') { $view = $Matches[1] }
    $docMode = 0; if ($pageInfo -match ' mode=(\d+)') { $docMode = [int]$Matches[1] }
    $err = ''; if ($pageInfo -match 'ERR\{(.*)\}') { $err = $Matches[1] }
    $attach = 0; if ($status -match ' mode=(\d) ') { $attach = [int]$Matches[1] }
    $k = 0.0; if ($ie -and $view -match '^(\d+)x') { $k = $ie.w / [double]$Matches[1] }
    if ($modern -and $ie) { $k = 1.0 }     # the developer protocol takes CSS px of the page
    # how smoothly the page appears: the entrance is played once more with the frame counter running
    $appear = $null; $act = $null
    if ($perf -and $ie) {
      if ((PageEval $t.page 'window.__ihFps ? window.__ihFps(1) : ""') -eq 'started') {
        [void](PageEval $t.page 'typeof window.pptHtmlOnShow === "function" ? (window.pptHtmlOnShow(), "ok") : "none"')
        Start-Sleep -Milliseconds 1700
        $appear = ReadFps $t.page
      }
    }
    # test actions for this slide, in the order given (canvas coordinates); the picture is taken after the last one
    $actions = @($A.click | Where-Object { $_ -and [int]$_.slide -eq $t.slide })
    $clicks = 0; $evals = @()
    if ($actions.Count -gt 0 -and $ie -and $k -gt 0) {
      if ($perf) { [void](PageEval $t.page 'window.__ihFps ? window.__ihFps(1) : ""') }
      foreach ($c in $actions) {
        $wait = 900; if ($c.wait -ne $null) { $wait = [int]$c.wait }
        $x = 0; $y = 0
        if ($c.x -ne $null) { $x = [int]([double]$c.x * $k); $y = [int]([double]$c.y * $k) }
        if ($c.eval) { $evals += ("{0} => {1}" -f [string]$c.eval, (PageEval $t.page ([string]$c.eval))) }
        elseif ($modern) {
          if ($c.wheel -ne $null) { MWheel $t.page ([double]$c.x) ([double]$c.y) ([double]$c.wheel) }
          elseif ($c.text -ne $null) { MText $t.page ([string]$c.text) }
          elseif ($c.key -ne $null) { MKey $t.page ([int]$c.key) }
          elseif ($c.to) { $hold = 60; if ($c.hold -ne $null) { $hold = [int]$c.hold }; MDrag $t.page ([double]$c.x) ([double]$c.y) ([double]@($c.to)[0]) ([double]@($c.to)[1]) 18 16 $hold }
          elseif ($c.hold -ne $null) { MClick $t.page ([double]$c.x) ([double]$c.y) ([int]$c.hold) }
          else { MClick $t.page ([double]$c.x) ([double]$c.y) 60 }
        }
        elseif ($c.text -ne $null) { [void][DeskKit.Desk]::PostText($script:HdDesk, $ie.h, [string]$c.text) }
        elseif ($c.key -ne $null) { [void][DeskKit.Desk]::PostKey($script:HdDesk, $ie.h, [int]$c.key) }
        elseif ($c.to) {
          $hold = 60; if ($c.hold -ne $null) { $hold = [int]$c.hold }
          [void][DeskKit.Desk]::PostDragSlow($script:HdDesk, $ie.h, $x, $y, [int]([double]@($c.to)[0] * $k), [int]([double]@($c.to)[1] * $k), 18, 16, $hold)
        }
        elseif ($c.hold -ne $null) { [void][DeskKit.Desk]::PostHold($script:HdDesk, $ie.h, $x, $y, [int]$c.hold) }
        else { [void][DeskKit.Desk]::PostClick($script:HdDesk, $ie.h, $x, $y) }
        Start-Sleep -Milliseconds $wait
        $clicks++
      }
      if ($perf) { $act = ReadFps $t.page }
      # script errors raised by the handlers show up in the status only now
      $status = ((Invoke-HdMacro $app $pres 'InlineHtml_Status') -split "`r`n")[0]
      if ($status -match (' p' + $t.page + '\[([^\]]*)\]')) { $pageInfo = $Matches[1] }
      $err = ''; if ($pageInfo -match 'ERR\{(.*)\}') { $err = $Matches[1] }
    }
    $shot = Join-Path ([string]$A.shotsDir) ("show_slide_{0:D2}.png" -f $t.slide)
    $shotOk = $false
    try { $shotOk = [bool](Save-HdShot $sw.Hwnd $shot 2); [void](Shrink $shot $maxW) } catch { L "shot failed: $($_.Exception.Message)" }
    $live = ([bool]$ie -and $pageInfo -match '^alive,hooked')
    # the browser's own picture of the page (the modern engine draws in a process of its own)
    $own = ''
    if ($modern -and $live) {
      $own = Join-Path ([string]$A.shotsDir) ("page_slide_{0:D2}.png" -f $t.slide)
      try { if ((Invoke-HdMacro $app $pres 'InlineHtml_Snap' @($t.page, $own)) -eq 'ok') { [void](Shrink $own $maxW) } else { $own = '' } } catch { $own = ''; L "own picture failed: $($_.Exception.Message)" }
    }
    # what the page has stored in the presentation so far (edits, kept values): characters in its slide tags
    $stateChars = 0
    for ($q = 0; $q -lt 200; $q++) { $piece = ''; try { $piece = [string]$pres.Slides.Item($t.slide).Tags.Item('HtmlState' + $q) } catch {}; if (-not $piece) { break }; $stateChars += $piece.Length }
    $results += [pscustomobject][ordered]@{ slide = $t.slide; page = $t.page; live = $live; engine = $(if ($modern) { 'modern' } else { 'classic' }); pagePicture = $own; view = $view; docMode = $docMode; attachMode = $attach; errors = $err; clicks = $clicks
      appearFps = $(if ($appear) { $appear.fps } else { 0 }); appearWorstMs = $(if ($appear) { $appear.worst } else { 0 })
      actionFps = $(if ($act) { $act.fps } else { 0 }); actionWorstMs = $(if ($act) { $act.worst } else { 0 })
      evals = @($evals); stateChars = $stateChars; shot = $(if ($shotOk) { $shot } else { '' }) }
    L ("slide {0} (page {1}): live={2} view={3} docMode={4} attach={5} errors=[{6}] appear={7}fps/{8}ms actions={9}fps/{10}ms" -f $t.slide, $t.page, $live, $view, $docMode, $attach, $err, $(if ($appear) { $appear.fps } else { '-' }), $(if ($appear) { $appear.worst } else { '-' }), $(if ($act) { $act.fps } else { '-' }), $(if ($act) { $act.worst } else { '-' }))
  }
  $full = Invoke-HdMacro $app $pres 'InlineHtml_Status'
  $engineNow = [string](Invoke-HdMacro $app $pres 'InlineHtml_Engine')
  $engineWhy = ''; if ($full -match 'engine classic: ([^\r\n]*)') { $engineWhy = $Matches[1].Trim() }
  $w.View.Exit(); $w = $null
  Start-Sleep -Milliseconds 1200
  # the host may still be drawing the stills of pages that stored something: wait until it has gone
  $endClock = [Diagnostics.Stopwatch]::StartNew(); $endStatus = ''
  while ($endClock.Elapsed.TotalSeconds -lt (20 + 6 * $targets.Count)) {
    try { $endStatus = [string](Invoke-HdMacro $app $pres 'InlineHtml_Status') } catch { $endStatus = '' }
    if ($endStatus -match '^pages=\d+ up=False ') { break }
    Start-Sleep -Milliseconds 300
  }
  $stillLog = @(($endStatus -split "`r`n") | Where-Object { $_ -match ' still' }) -join ' | '
  if ($stillLog) { L "stills: $stillLog" }
  if ([bool]$A.stills) {
    foreach ($t in $targets) {
      $sf = Join-Path ([string]$A.shotsDir) ("still_after_{0:D2}.png" -f $t.slide)
      try { $pres.Slides.Item($t.slide).Export($sf, 'PNG', 1600, [int](1600 * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)) } catch { L "still export failed: $($_.Exception.Message)" }
    }
  }
  if ($keep) { $pres.Save(); L 'saved (keep)' }
  # only the pages of this (hidden) PowerPoint: another PowerPoint of the user may be running a show of its own
  $myPid = 0; try { $myPid = [int](Get-HdPid) } catch {}
  $pat = 'ppthtml_*.html'; if ($myPid) { $pat = 'ppthtml_' + $myPid + '_*.html' }
  $left = @(Get-ChildItem -Path (Join-Path $env:TEMP $pat) -ErrorAction SilentlyContinue).Count
  if ($myPid) { $left += @(Get-ChildItem -Path (Join-Path $env:TEMP ('ppthtml_' + $myPid + '_m')) -File -ErrorAction SilentlyContinue).Count }
  $after = ((Invoke-HdMacro $app $pres 'InlineHtml_Status') -split "`r`n")[0]
  L "show window: $showDesc"
  L "status log: $full"
  L "after the show: $after ; temp pages left: $left"
  Write-JsonFile $A.result ([ordered]@{ engine = $engineNow; engineWhy = $engineWhy; browser = $browser; slides = @($results); showWindow = $showDesc; statusLog = $full; tempLeft = $left; firstLiveSec = $firstLive; stills = $stillLog })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($app.SlideShowWindows.Count -gt 0) { $app.SlideShowWindows.Item(1).View.Exit(); Start-Sleep -Milliseconds 800 } } catch {}
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
