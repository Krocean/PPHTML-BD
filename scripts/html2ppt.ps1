param(
  [string]$Job,            # job file (JSON, UTF-8): see references\job-file.md. Use it for anything beyond the simplest case.
  [string]$Html,           # simple form: one HTML file, or several separated by | (each becomes one full slide)
  [string]$Deck,           # existing presentation to put pages on (places marked {{html:file.html}}, or pages from the job file)
  [string]$Out,            # result (.pptm). Default: next to the deck / first HTML file.
  [string]$HtmlDir,        # folder that {{html:...}} markers and relative html paths refer to (default: the deck's / job file's folder)
  [string]$SlideSize = '16:9',   # for a new presentation: 16:9, 4:3 or WIDTHxHEIGHT in points
  [switch]$Web,            # also write the single-file web version (opens in any browser, no macros)
  [switch]$Share,          # also make a share folder + zip: .pptm, web version, a short read-me
  [switch]$Force,          # build even though the checks found problems
  [switch]$NoTest,         # skip the slide show test
  [switch]$KeepHidden,     # leave the hidden PowerPoint running afterwards
  [int]$SettleMs = -1,     # wait before the still picture of a page is taken (default 1500; job file: settleMs)
  [string]$Engine = '',    # which engine shows the pages in the slide show (job file: engine):
                           #   auto (default)  the modern engine (Edge WebView2): pages run exactly as written; plus, where the page allows it,
                           #                   a converted copy for the classic engine (Internet Explorer 11) for PCs without WebView2
                           #   modern          the modern engine only (faster build; PCs without WebView2 show the still pictures)
                           #   classic         the classic engine only (as before version 3: pages are converted for IE11 and checked against its rules)
  [switch]$TestFallback,   # engine auto: run the slide show test a second time with the classic engine (job file: testFallback)
  [string]$Convert = '',   # classic engine: auto (default): pages that use modern CSS / JavaScript are converted for IE11 automatically; always; never (job file: convert)
  [string]$Fonts = '',     # auto (default): the fonts a page uses are embedded in it; off (job file: fonts)
  [string]$Images = ''     # auto (default): large pictures are scaled down / stored as JPEG while they are pulled in; keep (job file: images)
)
# HTML -> PowerPoint: puts HTML pages on slides so that they are live (clickable, animated) during the slide show.
# The result is one self-contained .pptm; nothing is installed and no server runs. In the show the pages run in
# the engine of Microsoft Edge (WebView2, part of Windows 11 and of current Windows 10 / Office), so a page may
# use whatever a current browser offers; the classic engine (Internet Explorer 11) is the fallback.
# Exit codes: 0 done; 2 the pages need fixing (see the report); 3 the presentation / this PC cannot be used as
# asked; 4 built, but the slide show test did not see the live page; 1 unexpected error.
$ErrorActionPreference = 'Stop'
$PictureScale = 2     # still pictures are rendered at this multiple of the canvas (job file: pictureScale, 1-4)
$ShotWidth = 1600    # slide show test pictures are shrunk to this width (job file: shotWidth; 8000 keeps the original size)
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\pages.ps1')
. (Join-Path $PSScriptRoot 'lib\modern.ps1')
Initialize-Console
$exit = 0
$hdStarted = $false
$work = New-WorkDir 'html2ppt'
$Res = [ordered]@{ ok = $false; out = ''; web = ''; webCheck = $null; shareFolder = ''; shareZip = ''; workDir = $work; engine = ''; pages = @(); notes = @(); warnings = @() }
function Note([string]$s) { $script:Res.notes += $s; Say "  $s" }
function Warn([string]$s) { $script:Res.warnings += $s; Say "  WARNING: $s" }
function Fail([int]$code, [string]$msg) { $script:exit = $code; throw $msg }

try {
  $strings = Read-JsonFile (Join-Path $script:HdSkill 'assets\strings.json')

  # ---------------------------------------------------------------- what was asked
  $base = (Get-Location).Path
  $jobObj = $null
  if ($Job) {
    $jobPath = Resolve-UserPath $Job $base
    if (-not (Test-Path -LiteralPath $jobPath)) { Fail 3 "Job file not found: $jobPath" }
    $jobObj = Read-JsonFile $jobPath
    $base = Split-Path -Parent $jobPath
    if (-not $Deck -and $jobObj.deck) { $Deck = [string]$jobObj.deck }
    if (-not $Out -and $jobObj.out) { $Out = [string]$jobObj.out }
    if (-not $HtmlDir -and $jobObj.htmlDir) { $HtmlDir = [string]$jobObj.htmlDir }
    if ($jobObj.slideSize) { $SlideSize = [string]$jobObj.slideSize }
    if ($SettleMs -lt 0 -and $jobObj.settleMs -ne $null) { $SettleMs = [int]$jobObj.settleMs }
    if ($jobObj.pictureScale) { $PictureScale = [Math]::Max(1, [Math]::Min(4, [int]$jobObj.pictureScale)) }
    if ($jobObj.shotWidth -ne $null) { $ShotWidth = [int]$jobObj.shotWidth }
    if (-not $Convert -and $jobObj.convert -ne $null) { $Convert = [string]$jobObj.convert }
    if (-not $Fonts -and $jobObj.fonts -ne $null) { $Fonts = [string]$jobObj.fonts }
    if (-not $Images -and $jobObj.images -ne $null) { $Images = [string]$jobObj.images }
    if (-not $Engine -and $jobObj.engine -ne $null) { $Engine = [string]$jobObj.engine }
    if ($jobObj.testFallback) { $TestFallback = $true }
    if ($jobObj.web) { $Web = $true }
    if ($jobObj.share) { $Share = $true }
  }
  if ($SettleMs -lt 0) { $SettleMs = 1500 }
  if (-not $Convert) { $Convert = 'auto' }; if ($Convert -eq 'True') { $Convert = 'auto' }; if ($Convert -eq 'False') { $Convert = 'never' }
  if (-not $Images) { $Images = 'auto' }; if (-not $Fonts) { $Fonts = 'auto' }; if ($Fonts -eq 'True') { $Fonts = 'auto' }; if ($Fonts -eq 'False') { $Fonts = 'off' }
  if (@('auto', 'always', 'never') -notcontains $Convert.ToLowerInvariant()) { Fail 3 "Unknown -Convert value '$Convert' (use auto, always or never)." }
  if (-not $Engine) { $Engine = 'auto' }; $Engine = $Engine.ToLowerInvariant()
  if ($Engine -eq 'webview2' -or $Engine -eq 'edge') { $Engine = 'modern' }; if ($Engine -eq 'ie11' -or $Engine -eq 'ie') { $Engine = 'classic' }
  if (@('auto', 'modern', 'classic') -notcontains $Engine) { Fail 3 "Unknown -Engine value '$Engine' (use auto, modern or classic)." }
  if ($Share) { $Web = $true }
  $deckPath = Resolve-UserPath $Deck $base
  if ($deckPath -and -not (Test-Path -LiteralPath $deckPath)) { Fail 3 "Presentation not found: $deckPath" }
  $htmlBase = if ($HtmlDir) { Resolve-UserPath $HtmlDir $base } elseif ($jobObj) { $base } elseif ($deckPath) { Split-Path -Parent $deckPath } else { $base }

  $asked = @()    # { html, slide (int, 0 = new), area }
  if ($jobObj -and $jobObj.pages) {
    foreach ($p in @($jobObj.pages)) {
      if (-not $p.html) { Fail 3 'Job file: every page needs "html".' }
      $slide = 0
      if ($p.slide -ne $null -and "$($p.slide)" -ne 'new') { $slide = [int]$p.slide }
      $asked += [pscustomobject]@{ html = (Resolve-UserPath ([string]$p.html) $htmlBase); slide = $slide; area = $p.area }
    }
  }
  if ($Html) {
    foreach ($h in ($Html -split '\|')) { if ($h.Trim()) { $asked += [pscustomobject]@{ html = (Resolve-UserPath $h $base); slide = 0; area = $null } } }
  }

  # ---------------------------------------------------------------- the presentation
  $info = $null
  $deckCopy = ''
  if ($deckPath) {
    Say "[1/6] Reading the presentation (hidden PowerPoint) ..."
    Start-Hd | Out-Null; $hdStarted = $true
    $deckCopy = Join-Path $work ('deck_in' + [IO.Path]::GetExtension($deckPath).ToLowerInvariant())
    $run = Invoke-HdChild -Script 'deck_stage.ps1' -WorkDir $work -Params @{ deck = $deckPath; copyTo = $deckCopy }
    if (-not $run.Ok) { Fail 3 (Format-ChildFailure $run 'Reading the presentation') }
    $info = $run.Result
    Note "$($info.name): $($info.slideCount) slides, $($info.slideWidth) x $($info.slideHeight) pt"
    if ($info.hasVba -and -not $info.inlineDeck) { Fail 3 'This presentation already contains macros of its own. Adding the HTML host would replace them, so nothing was changed. Use a copy without macros.' }
  } else {
    Say "[1/6] New presentation ..."
    $sw = 960.0; $sh = 540.0
    if ($SlideSize -eq '4:3') { $sw = 720.0; $sh = 540.0 }
    elseif ($SlideSize -match '^\s*(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)\s*$') { $sw = [double]$Matches[1]; $sh = [double]$Matches[2] }
    elseif ($SlideSize -ne '16:9') { Fail 3 "Unknown slide size '$SlideSize' (use 16:9, 4:3 or WIDTHxHEIGHT in points)." }
    $info = [pscustomobject]@{ name = ''; slideWidth = $sw; slideHeight = $sh; slideCount = 0; hasVba = $false; inlineDeck = $false; slides = @() }
  }
  $SW = [double]$info.slideWidth; $SH = [double]$info.slideHeight

  # ---------------------------------------------------------------- pages and their places
  function Resolve-Rect($spec) {
    $one = {
      param($v, [double]$total)
      if ($v -is [string] -and $v.Trim().EndsWith('%')) { return ([double]$v.Trim().TrimEnd('%')) / 100.0 * $total }
      return [double]$v
    }
    $l = & $one $spec.left $SW; $t = & $one $spec.top $SH; $w = & $one $spec.width $SW; $h = & $one $spec.height $SH
    if ($l -lt 0) { $l = 0 }; if ($t -lt 0) { $t = 0 }
    if ($l + $w -gt $SW) { $w = $SW - $l }; if ($t + $h -gt $SH) { $h = $SH - $t }
    if ($w -lt 40 -or $h -lt 40) { throw 'an area must be at least 40 x 40 pt and lie on the slide' }
    return @{ left = $l; top = $t; width = $w; height = $h }
  }
  $pages = @()
  $taken = @{}
  foreach ($a in $asked) {
    $pg = [ordered]@{ html = $a.html; slide = $a.slide; shapeId = 0; left = 0.0; top = 0.0; width = $SW; height = $SH; where = 'full slide' }
    if ($a.slide -gt 0) {
      if ($a.slide -gt [int]$info.slideCount) { Fail 3 "Page $($a.html): slide $($a.slide) does not exist (the presentation has $($info.slideCount) slides)." }
      if ($taken.ContainsKey($a.slide)) { Fail 3 "Slide $($a.slide) was given two pages; one page per slide." }
      $taken[$a.slide] = $true
      $sl = @($info.slides)[$a.slide - 1]
      $spot = @($sl.areas) | Select-Object -First 1
      $explicit = ($a.area -ne $null -and -not ($a.area -is [string]))
      if ($explicit) {
        try { $r = Resolve-Rect $a.area } catch { Fail 3 "Page $($a.html): $($_.Exception.Message)" }
        $pg.left = $r.left; $pg.top = $r.top; $pg.width = $r.width; $pg.height = $r.height; $pg.where = 'given area'
        if ($spot) { $pg.shapeId = [int]$spot.id }
      } elseif ($spot -and "$($a.area)" -ne 'full') {
        $pg.shapeId = [int]$spot.id; $pg.left = [double]$spot.left; $pg.top = [double]$spot.top; $pg.width = [double]$spot.width; $pg.height = [double]$spot.height
        $pg.where = if ($spot.kind -eq 'marker') { 'marked shape' } else { 'area of the earlier build' }
      } elseif ("$($a.area)" -eq 'marker') { Fail 3 "Page $($a.html): slide $($a.slide) has no shape marked {{html:...}}." }
    } elseif ($a.area -ne $null -and -not ($a.area -is [string])) {
      try { $r = Resolve-Rect $a.area } catch { Fail 3 "Page $($a.html): $($_.Exception.Message)" }
      $pg.left = $r.left; $pg.top = $r.top; $pg.width = $r.width; $pg.height = $r.height; $pg.where = 'given area'
    }
    $pages += [pscustomobject]$pg
  }
  # shapes marked {{html:file}} on slides that were not given a page explicitly
  foreach ($sl in @($info.slides)) {
    if ($taken.ContainsKey([int]$sl.index)) { continue }
    $spot = @($sl.areas | Where-Object { $_.kind -eq 'marker' }) | Select-Object -First 1
    if (-not $spot) {
      if (@($sl.areas | Where-Object { $_.kind -eq 'area' }).Count -gt 0) { Warn "Slide $($sl.index) had a live page in the earlier build but none was given now: it keeps its still picture only." }
      continue
    }
    if (@($sl.areas | Where-Object { $_.kind -eq 'marker' }).Count -gt 1) { Warn "Slide $($sl.index) has more than one marked shape; only the first is used (one page per slide)." }
    $taken[[int]$sl.index] = $true
    $pages += [pscustomobject][ordered]@{ html = (Resolve-UserPath ([string]$spot.file) $htmlBase); slide = [int]$sl.index; shapeId = [int]$spot.id; left = [double]$spot.left; top = [double]$spot.top; width = [double]$spot.width; height = [double]$spot.height; where = 'marked shape' }
  }
  if ($pages.Count -eq 0) { Fail 3 'No pages to place: give -Html, a job file with "pages", or mark shapes in the presentation with {{html:file.html}}.' }
  if ($pages.Count -gt 30) { Fail 3 "Too many pages ($($pages.Count)); the limit is 30 per presentation." }
  if ($pages.Count -gt 12) { Warn "$($pages.Count) pages: each one is a browser of its own during the show; starting the show will take a few seconds." }
  $pages = @(@($pages | Where-Object { $_.slide -gt 0 } | Sort-Object slide) + @($pages | Where-Object { $_.slide -eq 0 }))
  $n = 0
  foreach ($p in $pages) {
    $n++
    $p | Add-Member -NotePropertyName n -NotePropertyValue $n
    $p | Add-Member -NotePropertyName canvasW -NotePropertyValue ([int][Math]::Round($p.width * 4 / 3))
    $p | Add-Member -NotePropertyName canvasH -NotePropertyValue ([int][Math]::Round($p.height * 4 / 3))
    if (-not (Test-Path -LiteralPath $p.html)) { Fail 3 "HTML file not found: $($p.html)" }
  }

  # ---------------------------------------------------------------- prepare + check the pages
  foreach ($p in $pages) { $p | Add-Member -NotePropertyName label -NotePropertyValue ("slide {0}, {1}" -f $(if ($p.slide) { $p.slide } else { 'new' }), $p.where) }
  $problem = $false
  if ($Engine -ne 'classic') {
    Say "[2/6] Preparing $($pages.Count) page(s) for the modern engine (Edge WebView2: pages run as written; fonts: $Fonts) ..."
    $mRun = Invoke-ModernPrep -Pages $pages -WorkDir $work -SettleMs $SettleMs -Scale $PictureScale -Fonts $Fonts -Images $Images
    if ($mRun.Fatal) { Fail $mRun.FatalCode $mRun.Fatal }
    if (-not $mRun.Ok) {
      if ($Engine -eq 'modern') { Fail 3 "The modern engine was asked for, but Edge cannot be used as the build tool on this PC: $($mRun.Why)" }
      Warn "Edge cannot be used as the build tool on this PC ($($mRun.Why)), so the pages are built for the classic engine (Internet Explorer 11) only."
      $Engine = 'classic'
    } else {
      Say $mRun.Report
      $problem = $mRun.Problem
      if ($problem -and -not $Force) {
        Say ''
        Say 'The pages need fixing before they can go on the slides (ERROR lines above): a file the page uses is missing, or the page''s own script fails in a current browser.'
        Fail 2 'pages need fixing'
      }
      foreach ($p in $pages) { if ($p.inlinedBytesM -gt 6MB) { Warn "Page $($p.n) pulls in $([Math]::Round($p.inlinedBytesM / 1MB, 1)) MB of files; the presentation grows by about that much." } }
    }
  }
  if ($Engine -eq 'classic') {
    Say "[2/6] Preparing $($pages.Count) page(s) for the classic engine (IE11; automatic conversion: $Convert, fonts: $Fonts) and rendering them off screen ..."
    $prepRun = Invoke-PagePrep -Pages $pages -WorkDir $work -SettleMs $SettleMs -Scale $PictureScale -Convert $Convert -Fonts $Fonts -Images $Images
    if ($prepRun.Fatal) { Fail $prepRun.FatalCode $prepRun.Fatal }
    foreach ($x in $prepRun.Warnings) { Warn $x }
    $problem = $prepRun.Problem
    Say $prepRun.Report
    if ($problem -and -not $Force) {
      Say ''
      Say 'The pages need fixing before they can go on the slides (ERROR lines above).'
      Say 'These are the things the automatic conversion for the classic engine could not do. Fix them in the HTML source files (references\ie11-rules.md has the rewrites), or build for the modern engine (-Engine auto), which runs pages as written.'
      Fail 2 'pages need fixing'
    }
    foreach ($p in $pages) { $p | Add-Member -NotePropertyName docM -NotePropertyValue '' -Force }
  } elseif ($Engine -eq 'auto') {
    # The same pages once more for the classic engine: the fallback on PCs without WebView2. It never blocks the
    # build: a page the classic engine cannot run simply has no fallback (its slide shows the still picture there).
    Say "      ... and for the classic engine (IE11), as the fallback for PCs without WebView2 ..."
    $cwork = Join-Path $work 'classic'; New-Item -ItemType Directory -Force $cwork | Out-Null
    $cRun = $null; $cWhy = ''
    try { $cRun = Invoke-PagePrep -Pages $pages -WorkDir $cwork -SettleMs $SettleMs -Scale $PictureScale -Convert $Convert -Fonts $Fonts -Images $Images } catch { $cWhy = $_.Exception.Message }
    if ($cRun -and $cRun.Fatal) { $cWhy = [string]$cRun.Fatal; $cRun = $null }
    $withFallback = 0
    foreach ($p in $pages) {
      $ok = ($cRun -ne $null -and $p.PSObject.Properties['doc'] -and $p.doc -and -not $p.prepProblem -and -not $p.renderProblem)
      if ($ok) {
        $withFallback++
        $like = ''
        try {
          if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
          $cmp = Join-Path $cwork ("pages\page_{0:D2}.fallback.png" -f $p.n)
          $f = ([ImgKit.Compare]::Run([string]$p.pngM, [string]$p.png, $cmp, 2600)) -split '\|'
          $like = ", its picture is $($f[0])% like the modern one (side by side: $cmp)"
        } catch {}
        Say ("      page {0}: fallback ready{1}{2}" -f $p.n, $(if ($p.converted) { ' (converted for IE11)' } else { '' }), $like)
      } else {
        $p | Add-Member -NotePropertyName doc -NotePropertyValue '' -Force
        $why = $cWhy
        if (-not $why -and $p.PSObject.Properties['reportLines']) { $why = (@(@($p.reportLines) + @($p.renderLines) | Where-Object { $_ -match '^\s+ERROR' } | ForEach-Object { ($_ -replace '^\s+ERROR\s*', '') -replace '\s+', ' ' } | Select-Object -First 2) -join '; ') }
        if ($why.Length -gt 220) { $why = $why.Substring(0, 220) + ' ...' }
        Say ("      page {0}: no fallback (a PC without WebView2 shows this slide's still picture): {1}" -f $p.n, $why)
      }
      $p | Add-Member -NotePropertyName classicPng -NotePropertyValue $(if ($p.PSObject.Properties['png']) { [string]$p.png } else { '' }) -Force
    }
    if ($cRun -and $cRun.ReportFile) { Say "      full report of the classic engine: $($cRun.ReportFile)" }
    Note "$withFallback of $($pages.Count) page(s) also run in the classic engine."
  } else {
    foreach ($p in $pages) { $p | Add-Member -NotePropertyName doc -NotePropertyValue '' -Force }
  }
  # the still picture and the background of each slide come from the engine the show will use
  if ($Engine -ne 'classic') {
    foreach ($p in $pages) {
      $p | Add-Member -NotePropertyName png -NotePropertyValue ([string]$p.pngM) -Force
      $p | Add-Member -NotePropertyName bg -NotePropertyValue ([string]$p.bgM) -Force
    }
  }
  if ($problem) { Warn 'Problems were found but -Force was given: building anyway.' }

  # ---------------------------------------------------------------- deck manifest
  # Every page is told the slide list (number, title, small picture of the page): its slide picker uses it.
  try {
    if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
    $total = [int]$info.slideCount + @($pages | Where-Object { $_.slide -eq 0 }).Count
    $slideOf = @{}; $nextNew = [int]$info.slideCount
    foreach ($p in $pages) { if ($p.slide -gt 0) { $slideOf[[int]$p.n] = [int]$p.slide } else { $nextNew++; $slideOf[[int]$p.n] = $nextNew } }
    $thumbW = if ($total -gt 16) { 168 } else { 224 }
    $list = New-Object Collections.ArrayList
    for ($i = 1; $i -le $total; $i++) {
      $entry = New-Object 'Collections.Generic.Dictionary[string,object]'
      $entry['n'] = $i; $entry['title'] = ''
      $pg = $pages | Where-Object { $slideOf[[int]$_.n] -eq $i } | Select-Object -First 1
      if ($pg) {
        $tdoc = if ($pg.docM) { [string]$pg.docM } else { [string]$pg.doc }
        if ($tdoc -match '<title[^>]*>([^<]*)</title>') { $entry['title'] = [Net.WebUtility]::HtmlDecode($Matches[1].Trim()) }
        if ($total -le 40 -and $pg.png -and (Test-Path -LiteralPath $pg.png)) { try { $entry['thumb'] = 'data:image/jpeg;base64,' + [ImgKit.Compare]::Thumb([string]$pg.png, $thumbW, 62) } catch {} }
      } elseif ($i -le @($info.slides).Count) { $entry['title'] = [string]@($info.slides)[$i - 1].title }
      [void]$list.Add($entry)
    }
    $deckId = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $mark = [PptHtml.HtmlPrep]::HeadClose
    foreach ($p in $pages) {
      $man = New-Object 'Collections.Generic.Dictionary[string,object]'
      $man['id'] = $deckId; $man['n'] = $slideOf[[int]$p.n]; $man['total'] = $total; $man['slides'] = $list.ToArray()
      $json = (ConvertTo-BigJson $man).Replace('</', '<\/')
      foreach ($which in 'doc', 'docM') {
        $text = [string]$p.$which
        if (-not $text) { continue }
        $at = $text.IndexOf($mark)
        if ($at -ge 0) { $p.$which = $text.Insert($at + $mark.Length, '<script>window.__ihDeck=' + $json + ';</script>') }
      }
    }
  } catch { Warn "The slide list could not be added to the pages (the slide picker will be empty): $($_.Exception.Message)" }

  # ---------------------------------------------------------------- build
  Say "[3/6] Putting the pages on the slides ..."
  if (-not $hdStarted) { Start-Hd | Out-Null; $hdStarted = $true }
  $xml = New-Object Text.StringBuilder
  [void]$xml.Append('<?xml version="1.0" encoding="UTF-8"?><inlineHtml xmlns="urn:ppt-html:inline:v1">')
  foreach ($p in $pages) {
    # html = the document for the classic engine, htmlm = the document for the modern engine (either may be empty)
    [void]$xml.Append("<page><bg>$(ConvertTo-OleColor $p.bg)</bg><w>$($p.canvasW)</w><h>$($p.canvasH)</h><settle>$([int]$SettleMs)</settle><html>")
    [void]$xml.Append([PptHtml.HtmlPrep]::XmlEscape([string]$p.doc))
    [void]$xml.Append('</html><htmlm>')
    [void]$xml.Append([PptHtml.HtmlPrep]::XmlEscape([string]$p.docM))
    [void]$xml.Append('</htmlm></page>')
  }
  [void]$xml.Append('</inlineHtml>')
  $xmlFile = Join-Path $work 'pages.xml'
  [IO.File]::WriteAllText($xmlFile, $xml.ToString(), $script:Utf8NoBom)
  $ap = @()
  foreach ($p in $pages) {
    $alt = ([string]$strings.areaAlt) -replace '\{file\}', [IO.Path]::GetFileName($p.html)
    $ap += @{ n = $p.n; slide = $p.slide; shapeId = $p.shapeId; left = $p.left; top = $p.top; width = $p.width; height = $p.height; png = $p.png; bg = $p.bg; alt = $alt }
  }
  $stage = Join-Path $work 'stage.pptm'
  $run = Invoke-HdChild -Script 'deck_apply.ps1' -WorkDir $work -Params @{ deck = $deckCopy; slideWidth = $SW; slideHeight = $SH; xmlFile = $xmlFile; stageOut = $stage; pages = $ap }
  if (-not $run.Ok) { Fail 3 (Format-ChildFailure $run 'Putting the pages on the slides') }
  foreach ($pl in @($run.Result.placed)) {
    ($pages | Where-Object { $_.n -eq $pl.n }).slide = [int]$pl.slide
    $cov = @($pl.covers | Where-Object { $_ })
    if ($cov.Count -gt 0) {
      # shapes that are only partly under the area come first: a cut-off title is the case that looks broken
      $cov = @(@($cov | Where-Object { $_ -like '*(partly covered)' }) + @($cov | Where-Object { $_ -notlike '*(partly covered)' }))
      $list = (@($cov | Select-Object -First 6) -join '; ') + $(if ($cov.Count -gt 6) { "; and $($cov.Count - 6) more" } else { '' })
      Warn "Slide $($pl.slide): the HTML area lies over existing slide content, which the live page hides or cuts during the show: $list. Move or shrink the area (or those shapes) unless that is intended."
    }
  }

  Say "[4/6] Adding the host code ..."
  $vbaBin = Join-Path $work 'vbaProject.bin'
  [void](New-VbaProject $vbaBin)
  $staged2 = Join-Path $work 'stage_vba.pptm'
  foreach ($x in (Add-VbaToPackage $stage $vbaBin $staged2)) { if ($x -like 'WARNING:*') { Warn $x.Substring(8).Trim() } else { Note $x } }
  $final = Join-Path $work 'final.pptm'
  $run = Invoke-HdChild -Script 'deck_finalize.ps1' -WorkDir $work -Params @{ file = $staged2; saveTo = $final }
  if (-not $run.Ok) { Fail 3 (Format-ChildFailure $run 'Finishing the presentation') }
  if (-not $run.Result.hasVba) { Fail 1 'PowerPoint did not accept the host code (the saved file has no macro project).' }

  $testFailed = $false
  if ($NoTest) { Say "[5/6] Slide show test skipped (-NoTest)." }
  else {
    $click = @(); if ($jobObj -and $jobObj.testClicks) { $click = @($jobObj.testClicks) }
    $passes = @('')     # '' = the engine the show picks by itself; 'classic' = the fallback, when asked for
    if ($TestFallback -and $Engine -eq 'auto') { $passes += 'classic' }
    foreach ($pass in $passes) {
    $isFallback = ($pass -eq 'classic')
    if ($isFallback) { Say "      ... and once more with the classic engine (the fallback) ..." } else { Say "[5/6] Testing in a real slide show (hidden) ..." }
    $shots = Join-Path $work $(if ($isFallback) { 'shots_classic' } else { 'shots' }); New-Item -ItemType Directory -Force $shots | Out-Null
    $run = Invoke-HdChild -Script 'show_test.ps1' -WorkDir $work -Params @{ file = $final; shotsDir = $shots; showType = 1; click = $click; maxShotWidth = $ShotWidth; engine = $pass; noClassic = @($pages | Where-Object { -not $_.doc } | ForEach-Object { [int]$_.n }) } -TimeoutSec (120 + 30 * $pages.Count)
    if (-not $run.Ok) { $testFailed = $true; Warn (Format-ChildFailure $run 'The slide show test') }
    else {
      $shown = [string]$run.Result.engine
      if (-not $isFallback) {
        $Res.engine = $shown
        if ($Engine -ne 'classic' -and $shown -ne 'modern') { Warn "The slide show on this PC did not use the modern engine but the classic one$(if ($run.Result.engineWhy) { " ($($run.Result.engineWhy))" }). Pages that have no fallback stayed still pictures." }
      }
      foreach ($t in @($run.Result.slides)) {
        $p = $pages | Where-Object { $_.n -eq $t.page }
        $noDoc = ($shown -ne 'modern' -and -not $p.doc)      # the classic engine has no document for this page
        if (-not $isFallback) {
          $p | Add-Member -NotePropertyName live -NotePropertyValue ([bool]$t.live) -Force
          $p | Add-Member -NotePropertyName shot -NotePropertyValue ([string]$t.shot) -Force
          $p | Add-Member -NotePropertyName showView -NotePropertyValue ([string]$t.view) -Force
        }
        $state = if ($t.live) { 'LIVE' } elseif ($noDoc) { 'STILL PICTURE (no fallback for this page)' } else { 'NOT LIVE' }
        $clickNote = if ([int]$t.clicks -gt 0) { ", picture taken after $($t.clicks) test action(s)" } else { '' }
        $fpsNote = ''
        if ([int]$t.appearFps -gt 0) { $fpsNote = ", $($t.appearFps) fps while appearing (longest frame $($t.appearWorstMs) ms)" }
        if ([int]$t.actionFps -gt 0) { $fpsNote += ", $($t.actionFps) fps during the test actions (longest frame $($t.actionWorstMs) ms)" }
        if (-not $isFallback) {
          $p | Add-Member -NotePropertyName appearFps -NotePropertyValue ([int]$t.appearFps) -Force
          $p | Add-Member -NotePropertyName actionFps -NotePropertyValue ([int]$t.actionFps) -Force
        }
        $engText = if ($shown -eq 'modern') { "modern engine$(if ($run.Result.browser) { " (Edge $($run.Result.browser))" })" } else { "classic engine (document mode $($t.docMode))" }
        Say ("  slide {0}: {1}  viewport {2} px, {3}{4}{5}  picture: {6}" -f $t.slide, $state, $t.view, $engText, $(if ($t.errors) { ", script errors: $($t.errors)" } else { '' }), ($clickNote + $fpsNote), $t.shot)
        if (-not $t.live -and -not $noDoc) { $testFailed = $true }
        if ($t.errors) { Warn "Slide $($t.slide): script error during the show: $($t.errors)" }
        foreach ($ev in @($t.evals)) { Say "      eval: $ev" }
        $slowest = @([int]$t.appearFps, [int]$t.actionFps | Where-Object { $_ -gt 0 } | Sort-Object | Select-Object -First 1)
        if ($slowest.Count -gt 0 -and $slowest[0] -lt 30) { Warn "Slide $($t.slide) moves at $($slowest[0]) frames per second in the slide show ($engText; below 30 looks jerky)." }
      }
      if ([int]$run.Result.tempLeft -gt 0) { Warn "$($run.Result.tempLeft) temporary page file(s) were left in %TEMP% after the show." }
      foreach ($d in $run.Dialogs) { Warn "PowerPoint showed a dialog during the test: $d" }
      if ($testFailed) { [IO.File]::WriteAllText((Join-Path $work 'show-status.txt'), [string]$run.Result.statusLog, $script:Utf8NoBom); Warn "The live page was not seen on every slide. Host log: $(Join-Path $work 'show-status.txt')" }
    }
    }
  }

  # ---------------------------------------------------------------- deliver
  Say "[6/6] Saving ..."
  if (-not $Out) {
    if ($deckPath) { $Out = Join-Path (Split-Path -Parent $deckPath) ([IO.Path]::GetFileNameWithoutExtension($deckPath) + '_HTML.pptm') }
    else { $Out = [IO.Path]::ChangeExtension($pages[0].html, '.pptm') }   # (a web version next to it gets the suffix _web: see below)
  }
  $outPath = Resolve-UserPath $Out $base
  if ([IO.Path]::GetExtension($outPath).ToLowerInvariant() -ne '.pptm') { $outPath = [IO.Path]::ChangeExtension($outPath, '.pptm'); Note 'The result has to be a .pptm (macro-enabled presentation); the extension was changed.' }
  $written = Copy-Deliver $final $outPath
  if (-not $written) {
    # PowerShell may not write there (protected folder): let PowerPoint save the copy
    $run = Invoke-HdChild -Script 'deck_savecopy.ps1' -WorkDir $work -Params @{ file = $final; saveTo = $outPath }
    if ($run.Ok) { $written = $outPath } else { Warn "Could not write $outPath. The finished file is here: $final" ; $written = $final }
  }
  if ($written -ne $outPath -and $written -ne $final) { Note "The name was in use (file open?), so a numbered name was used." }
  $Res.out = $written
  Say "  $written"

  if ($Web) {
    . (Join-Path $PSScriptRoot 'lib\webdeck.ps1')
    . (Join-Path $PSScriptRoot 'lib\webcheck.ps1')
    $webOut = [IO.Path]::ChangeExtension($written, '.html')
    # never write the web version over a page source that happens to have the same name
    if (@($pages | Where-Object { [string]::Equals([IO.Path]::GetFullPath([string]$_.html), [IO.Path]::GetFullPath($webOut), [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) {
      $webOut = Join-Path (Split-Path -Parent $webOut) ([IO.Path]::GetFileNameWithoutExtension($webOut) + '_web.html')
      Note 'The web version is named ..._web.html because a page source already has the plain name.'
    }
    $wd = New-WebDeck -Source $final -OutFile $webOut -WorkDir $work -Title ([IO.Path]::GetFileNameWithoutExtension($written))
    foreach ($x in $wd.warnings) { Warn $x }
    $Res.web = $wd.out
    Say "  $($wd.out)   (web version for any browser)"
    $shotsW = Join-Path $work 'shots'; if (-not (Test-Path -LiteralPath $shotsW)) { New-Item -ItemType Directory -Force $shotsW | Out-Null }
    $wc = Test-WebDeck -File $wd.out -WorkDir $work -ShotsDir $shotsW -SettleMs $SettleMs
    $Res.webCheck = [ordered]@{ checked = $wc.checked; slides = $wc.slides; livePages = $wc.live; liveOk = $wc.liveOk; errors = @($wc.errors); pictures = @($wc.shots) }
    if (-not $wc.checked) { Say "    web version not checked in a browser: $($wc.why)" }
    else {
      Say "    checked in Edge: $($wc.slides) slide(s), $($wc.liveOk) of $($wc.live) live page(s) running$(if (@($wc.errors).Count -eq 0) { ', no script errors' } else { '' })"
      foreach ($e in @($wc.errors)) { Warn "Web version: $e" }
      foreach ($s in @($wc.shots)) { Say "    picture: $s" }
    }
    if ($Share) {
      $sb = New-ShareBundle -Pptm $written -Web $wd.out -WorkDir $work -Strings $strings
      $Res.shareFolder = $sb.folder; $Res.shareZip = $sb.zip
      Say "  $($sb.zip)   (share bundle: the .pptm, the web version and a read-me)"
    }
  }

  foreach ($p in $pages) {
    $Res.pages += [ordered]@{ n = $p.n; html = $p.html; slide = $p.slide; area = @([Math]::Round($p.left, 1), [Math]::Round($p.top, 1), [Math]::Round($p.width, 1), [Math]::Round($p.height, 1)); canvas = @($p.canvasW, $p.canvasH); preview = $p.png; modern = [bool]$p.docM; fallback = [bool]$p.doc; scrolls = [bool]$p.scrolls; prepared = $(if ($p.docM) { $p.fileM } else { $p.readable }); converted = [bool]$p.converted; classicPicture = $p.classicPng; comparePicture = $p.comparePng; live = $p.live; showPicture = $p.shot; appearFps = $p.appearFps; actionFps = $p.actionFps; lintWarnings = $p.lintWarnings; lintErrors = $p.lintErrors }
  }
  $Res.ok = $true
  if ($testFailed) { $exit = 4 }
  Say ''
  if ($testFailed) { Say 'DONE, but the slide show test did not confirm every live page (see warnings).' }
  elseif ($NoTest) { Say 'DONE (not tested in a slide show).' }
  else { Say 'DONE. The .pptm was run as a real (hidden) slide show on this PC and every page was live: loaded, script running, no script errors. Look at the pictures to judge the layout.' }
} catch {
  if ($exit -eq 0) { $exit = 1 }
  $Res.error = $_.Exception.Message
  if ($exit -ne 2) { Say "FAILED: $($_.Exception.Message)" }
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
} finally {
  if ($hdStarted -and -not $KeepHidden) { try { [void](Stop-Hd) } catch {} }
  try { Write-JsonFile (Join-Path $work 'result.json') $Res } catch {}
  Say "Details: $(Join-Path $work 'result.json')"
}
exit $exit
