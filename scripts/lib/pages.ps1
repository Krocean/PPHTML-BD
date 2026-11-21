# Dot-source (after common.ps1). Prepares HTML pages for the IE11 engine and checks them: one self-contained
# document per page, modern CSS / JavaScript converted automatically (Edge as the build tool), fonts embedded,
# a report of what the engine still cannot do, a real render in the engine (picture + script errors).
. (Join-Path $PSScriptRoot 'edge.ps1')
. (Join-Path $PSScriptRoot 'fonts.ps1')

# Pages: objects with  html (source path), n (1..), canvasW, canvasH (CSS px), label (text for the report).
# Convert: auto (only pages that use what IE11 lacks) | always | never.   Fonts: auto (embed) | off.
# Images: auto (large pictures are scaled down / stored as JPEG, see HtmlPrep.Slim) | keep.
# Adds to each page: doc, file, readable, png, bg, lintErrors, lintWarnings, scriptErrors, converted, modernPng, comparePng.
# Returns: Problem (the pages need fixing), Report (text), Warnings, Fatal (message; nothing could be done), FatalCode.
function Invoke-PagePrep {
  param([Parameter(Mandatory)]$Pages, [Parameter(Mandatory)][string]$WorkDir, [int]$SettleMs = 1500, [int]$Scale = 2, [string]$Convert = 'auto', [string]$Fonts = 'auto', [string]$Probe = '', [string]$Images = 'auto')
  $ret = [pscustomobject]@{ Problem = $false; Report = ''; Warnings = @(); Fatal = $null; FatalCode = 0; ReportFile = ''; Converted = 0; Sheet = '' }
  Initialize-HtmlPrep
  $headH = Read-Asset 'harness-head.txt'; $tailPpt = Read-Asset 'harness-tail-ppt.txt'; $poly = Read-Asset 'polyfills-ie11.txt'
  $ctx = @{ heavy = $null; adapter = $null; fatal = $null; fatalCode = 0 }
  $pagesDir = Join-Path $WorkDir 'pages'
  if (-not (Test-Path -LiteralPath $pagesDir)) { New-Item -ItemType Directory -Force $pagesDir | Out-Null }
  $motw = '<!-- saved from url=(0014)about:internet -->'
  $denied = "ACCESS-DENIED: PowerShell is not allowed to read '{0}' on this PC (security software). Copy the page and the files it uses to a folder under '$(Join-Path (Get-HdRoot) 'in')' with your own file tools, then run again with that copy."
  $Convert = $Convert.ToLowerInvariant(); $Fonts = $Fonts.ToLowerInvariant()

  # one page: read, pull files in, convert when needed, embed fonts, add the harness, check against the rules
  $buildPage = {
    param($p, [bool]$force)
    $lines = New-Object 'Collections.Generic.List[string]'
    $src = $null; $st = $null; $prep = $null
    try { $src = [PptHtml.HtmlPrep]::ReadPage($p.html) }
    catch { $ctx.fatal = ($denied -f $p.html); $ctx.fatalCode = 3; return }
    # pictures: nothing on a slide needs more pixels than three times its canvas (a 4K screen at full size)
    [PptHtml.HtmlPrep]::ImageMaxSide = $(if (@('keep', 'off', 'never', 'false') -contains $Images.ToLowerInvariant()) { 0 } else { [Math]::Max(1920, 3 * [Math]::Max([int]$p.canvasW, [int]$p.canvasH)) })
    try { $st = [PptHtml.HtmlPrep]::Begin($src, (Split-Path -Parent $p.html)) }
    catch {
      $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
      if ($e -is [UnauthorizedAccessException]) { $ctx.fatal = ($denied -f "a file used by $($p.html): $($e.Message)"); $ctx.fatalCode = 3 }
      else { $ctx.fatal = "Preparing $($p.html) failed: $($e.Message)"; $ctx.fatalCode = 1 }
      return
    }
    $wantConv = (-not $st.OptOut) -and ($Convert -ne 'never') -and ($force -or $Convert -eq 'always' -or $st.NeedsJs -or $st.NeedsCss)
    $wantFonts = ($Fonts -ne 'off' -and $Fonts -ne 'never' -and $Fonts -ne 'false')
    $edge = $null
    $edgeAgain = '    note: the tool browser (Edge) stopped answering in the middle of the run and was started again'
    if ($wantConv -or $wantFonts) {
      $edge = Start-EdgeTool $WorkDir
      if ($edge -and -not (Test-EdgeTool)) { $edge = Restart-EdgeTool $WorkDir; if ($edge) { $lines.Add($edgeAgain) } }
    }
    $converted = $false; $modern = ''
    if ($wantConv) {
      if ($edge) {
        # a tool browser that went away is started again once and the step repeated (the page state only changes after Edge has answered)
        for ($try = 0; $try -lt 2 -and -not $converted; $try++) {
          try { [void](Convert-PageForIe $edge $st ([int]$p.canvasW) ([int]$p.canvasH)); $converted = $true }
          catch {
            $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
            if ($try -eq 0 -and -not (Test-EdgeTool)) { $edge = Restart-EdgeTool $WorkDir; if ($edge) { $lines.Add($edgeAgain); continue } }
            $lines.Add("    warn  the automatic conversion failed and was skipped: $($e.Message)"); break
          }
        }
      } else { $lines.Add("    note: no automatic conversion ($(Get-EdgeWhyNot)): the page has to be IE11-ready as written") }
    }
    # maintenance hook: PPTHTML_EDGE_DROP=<page number> closes the tool browser here, as if its connection had been cut
    if ($env:PPTHTML_EDGE_DROP -eq [string]$p.n -and $edge -and -not $ctx.dropped) { $ctx.dropped = $true; try { $edge.Dispose() } catch {} }
    if ($edge -and ($converted -or $wantFonts)) {
      for ($try = 0; $try -lt 2; $try++) {
        try {
          $before = $edge.Errors.Count
          $edge.SetViewport([int]$p.canvasW, [int]$p.canvasH, [double]$Scale)
          [void](Invoke-EdgeTool $edge 'IHTool.show' @(("<!DOCTYPE html>`n" + $st.Doc), [int]$p.canvasW, [int]$p.canvasH, $(if ($converted) { $SettleMs } else { 200 })) 60000)
          if ($converted) { $modern = Join-Path $pagesDir ("page_{0:D2}.modern.png" -f $p.n); $edge.Screenshot($modern) }
          if ($wantFonts) {
            $fontNotes = @(Add-PageFonts $edge $st)
            foreach ($x in $fontNotes) { $st.R.Notes.Add([string]$x) }
            # fonts that are not installed on this PC were just embedded: the reference picture is taken again with them,
            # so that the comparison shows differences of the layout and not of the fonts
            if ($converted -and $st.ExtraCss -and @($fontNotes | Where-Object { [string]$_ -like 'fonts embedded*' }).Count -gt 0) {
              $extra = [string]$st.ExtraCss
              $refDoc = (New-Object Text.RegularExpressions.Regex('<head\b[^>]*>', 'IgnoreCase')).Replace([string]$st.Doc, [Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Value + "`n<style>`n" + $extra + "</style>" }, 1)
              [void](Invoke-EdgeTool $edge 'IHTool.show' @(("<!DOCTYPE html>`n" + $refDoc), [int]$p.canvasW, [int]$p.canvasH, $SettleMs) 60000)
              $edge.Screenshot($modern)
            }
          }
          [void]$edge.Eval('IHTool.hide()', $false, 10000)
          if ($converted) { for ($k = $before; $k -lt $edge.Errors.Count -and $k -lt $before + 2; $k++) { $lines.Add("    note: even a modern browser reports an error for this page: $($edge.Errors[$k])") } }
          break
        } catch {
          $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
          if ($try -eq 0 -and -not (Test-EdgeTool)) { $edge = Restart-EdgeTool $WorkDir; if ($edge) { $lines.Add($edgeAgain); continue } }
          $lines.Add("    note: the modern-browser pass (reference picture, fonts) was skipped: $($e.Message)"); $modern = ''
          if ($edge) { try { [void]$edge.Eval('IHTool.hide()', $false, 5000) } catch {} }
          break
        }
      }
    } elseif ($wantFonts -and -not $edge) { $lines.Add("    note: fonts were not embedded ($(Get-EdgeWhyNot))") }
    $heavy = ''; $adapter = ''
    if ($converted) {
      if ($ctx.heavy -eq $null) { $ctx.heavy = Read-Asset 'corejs.min.txt'; $ctx.adapter = Read-Asset 'adapt-ie11.txt' }
      $heavy = $ctx.heavy; $adapter = $ctx.adapter
    }
    try { $prep = [PptHtml.HtmlPrep]::Finish($st, $headH, $tailPpt, $poly, $heavy, $adapter) }
    catch { $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }; $ctx.fatal = "Preparing $($p.html) failed: $($e.Message)"; $ctx.fatalCode = 1; return }
    $p | Add-Member -NotePropertyName doc -NotePropertyValue $prep.Html -Force
    $p | Add-Member -NotePropertyName file -NotePropertyValue (Join-Path $pagesDir ("page_{0:D2}.html" -f $p.n)) -Force
    $p | Add-Member -NotePropertyName readable -NotePropertyValue (Join-Path $pagesDir ("page_{0:D2}.prepared.txt" -f $p.n)) -Force
    $p | Add-Member -NotePropertyName png -NotePropertyValue (Join-Path $pagesDir ("page_{0:D2}.png" -f $p.n)) -Force
    $p | Add-Member -NotePropertyName converted -NotePropertyValue $converted -Force
    $p | Add-Member -NotePropertyName optOut -NotePropertyValue ([bool]$st.OptOut) -Force
    $p | Add-Member -NotePropertyName modernPng -NotePropertyValue $modern -Force
    $p | Add-Member -NotePropertyName comparePng -NotePropertyValue '' -Force
    # the file the engine loads is written exactly the way the host writes it during the show (CR LF, Mark of the Web)
    [IO.File]::WriteAllText($p.file, $prep.Html.Replace("`r`n", "`n").Replace("`n", "`r`n").Replace('<!DOCTYPE html>', "<!DOCTYPE html>`r`n$motw"), $script:Utf8NoBom)
    [IO.File]::WriteAllText($p.readable, $prep.Html, $script:Utf8NoBom)
    $errs = @($prep.Issues | Where-Object { $_.Severity -eq 'error' }); $warns = @($prep.Issues | Where-Object { $_.Severity -ne 'error' })
    $p | Add-Member -NotePropertyName lintErrors -NotePropertyValue $errs.Count -Force
    $p | Add-Member -NotePropertyName lintWarnings -NotePropertyValue $warns.Count -Force
    $problem = $false
    $head = "page {0}: {1}  ->  {2}, canvas {3} x {4} px{5}" -f $p.n, $p.html, $p.label, $p.canvasW, $p.canvasH, $(if ($converted) { '  [converted automatically for IE11]' } else { '' })
    $out = New-Object 'Collections.Generic.List[string]'
    $out.Add($head)
    foreach ($x in $prep.Notes) {
      if ($x -like 'converted: LOOK: *') { $out.Add('    look  ' + $x.Substring(17)) } else { $out.Add("    note: $x") }
    }
    foreach ($x in $lines) { $out.Add($x) }
    foreach ($x in $prep.Missing) { $out.Add("    ERROR missing file: $x"); $problem = $true }
    foreach ($x in $prep.External) { $out.Add("    warn  needs internet during the show: $x") }
    if (@($prep.Issues).Count -gt 0) { $out.Add("    (line numbers below refer to the prepared page: $($p.readable))") }
    foreach ($i in $prep.Issues) {
      $tag = if ($i.Severity -eq 'error') { 'ERROR' } else { 'warn ' }
      $where = if ($i.Lines) { " line $($i.Lines):" } elseif ($i.Line -gt 0) { " line $($i.Line):" } else { '' }
      $out.Add("    $tag [$($i.Rule)]$where $($i.Message)")
      if ($i.Snippet) { $out.Add("          > $($i.Snippet)") }
    }
    if ($errs.Count -gt 0) { $problem = $true }
    $p | Add-Member -NotePropertyName reportLines -NotePropertyValue $out -Force
    $p | Add-Member -NotePropertyName prepProblem -NotePropertyValue $problem -Force
    $p | Add-Member -NotePropertyName inlinedBytes -NotePropertyValue ([long]$prep.InlinedBytes) -Force
  }

  # render in the IE11 engine; adds the findings to the page's report lines
  $renderPages = {
    param($list)
    $rp = @(); foreach ($p in $list) { $rp += @{ n = $p.n; file = $p.file; canvasW = $p.canvasW; canvasH = $p.canvasH; png = $p.png } }
    $run = Invoke-HdChild -Script 'render.ps1' -WorkDir $WorkDir -NoHd -Params @{ pages = $rp; settleMs = $SettleMs; scale = $Scale; probe = $Probe } -TimeoutSec (120 + 30 * @($list).Count)
    if (-not $run.Ok) { $ctx.fatal = (Format-ChildFailure $run 'Rendering the pages'); $ctx.fatalCode = 1; return }
    foreach ($r in @($run.Result.pages)) {
      $p = $list | Where-Object { $_.n -eq $r.n }
      $p | Add-Member -NotePropertyName bg -NotePropertyValue ([string]$r.bg) -Force
      $p | Add-Member -NotePropertyName scriptErrors -NotePropertyValue ([string]$r.errors) -Force
      $bad = $false; $lines = @()
      if (-not $r.loaded) { $lines += 'ERROR the page did not finish loading within 20 s'; $bad = $true }
      if ($r.docMode -and $r.docMode -lt 11) { $lines += "ERROR the engine fell back to document mode $($r.docMode) (expected 11)"; $bad = $true }
      if ($r.errors) { $lines += "ERROR script error(s) while loading: $($r.errors)   (line numbers refer to $($p.readable))"; $bad = $true }
      elseif (-not $r.ready) { $lines += 'ERROR the page script did not run to the end'; $bad = $true }
      if ($r.clientW -and ($r.scrollW -gt $r.clientW + 2 -or $r.scrollH -gt $r.clientH + 2)) { $lines += "warn  content is $($r.scrollW) x $($r.scrollH) px but the canvas is $($r.clientW) x $($r.clientH) px: the rest is cut off or scrolls" }
      if ($r.perf -and ([string]$r.perf) -match '^(\d+)\|(\d+)\|(.*)$' -and [int]$Matches[1] -gt 0) {
        $pl = "warn  [perf-large-motion] $($Matches[1]) large element(s) fade or move (a transition or animation on something bigger than an eighth of the page): $($Matches[3]). The slide show engine repaints all of it on every frame, which shows as stutter; animate the small parts inside instead (references\ie11-rules.md, smooth motion)."
        if ([int]$Matches[2] -gt 0) { $pl += " $($Matches[2]) of them never stop (animation ... infinite)." }
        $lines += $pl
      }
      if ($Probe) { $lines += "probe: $($r.probe)" }
      $p | Add-Member -NotePropertyName renderProblem -NotePropertyValue $bad -Force
      $p | Add-Member -NotePropertyName scriptFailed -NotePropertyValue ([bool]($r.errors -or -not $r.ready)) -Force
      $p | Add-Member -NotePropertyName renderLines -NotePropertyValue (@("page $($p.n) rendered: document mode $($r.docMode), background $($r.bg), picture (${Scale}x the canvas): $($p.png)") + @($lines | ForEach-Object { "    $_" })) -Force
    }
  }

  try {
    foreach ($p in $Pages) { & $buildPage $p $false; if ($ctx.fatal) { $ret.Fatal = $ctx.fatal; $ret.FatalCode = $ctx.fatalCode; return $ret } }
    & $renderPages $Pages
    if ($ctx.fatal) { $ret.Fatal = $ctx.fatal; $ret.FatalCode = $ctx.fatalCode; return $ret }
    # second chance: a page that was not converted but whose script failed in the engine is converted and rendered again
    if ($Convert -eq 'auto') {
      $retry = @($Pages | Where-Object { $_.scriptFailed -and -not $_.converted -and -not $_.optOut })
      if ($retry.Count -gt 0 -and (Start-EdgeTool $WorkDir)) {
        foreach ($p in $retry) { & $buildPage $p $true; if ($ctx.fatal) { $ret.Fatal = $ctx.fatal; $ret.FatalCode = $ctx.fatalCode; return $ret } }
        $again = @($retry | Where-Object { $_.converted })
        if ($again.Count -gt 0) {
          foreach ($p in $again) { $p.reportLines.Insert(1, '    note: the script failed in the IE11 engine as written, so the page was converted automatically and rendered again') }
          & $renderPages $again
          if ($ctx.fatal) { $ret.Fatal = $ctx.fatal; $ret.FatalCode = $ctx.fatalCode; return $ret }
        }
      }
    }
    # converted pages: how close is the engine's picture to the modern browser's
    foreach ($p in $Pages) {
      if (-not $p.converted -or -not $p.modernPng -or -not (Test-Path -LiteralPath $p.modernPng) -or -not (Test-Path -LiteralPath $p.png)) { continue }
      try {
        if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
        $cmp = Join-Path $pagesDir ("page_{0:D2}.compare.png" -f $p.n)
        $f = ([ImgKit.Compare]::Run($p.modernPng, $p.png, $cmp, 2600)) -split '\|'
        $p.comparePng = $cmp
        $sim = [double]::Parse($f[0], [Globalization.CultureInfo]::InvariantCulture); $diff = [int]$f[1]
        $tag = if ($sim -lt 93 -or $diff -gt 22) { 'warn ' } else { 'look ' }
        $msg = "    $tag conversion check: the engine's picture is $($f[0])% like the modern browser's; $diff of $($f[2]) areas differ"
        if ($diff -gt 0) { $msg += " (most around column $($f[3]) of 16, row $($f[4]) of 9)" }
        $p.renderLines += "$msg. Side by side: $cmp"
      } catch { $p.renderLines += "    note: the comparison picture could not be made: $($_.Exception.Message)" }
    }
  } finally { Stop-EdgeTool }

  $report = New-Object Text.StringBuilder
  foreach ($p in $Pages) {
    foreach ($x in $p.reportLines) { [void]$report.AppendLine($x) }
    if ($p.prepProblem) { $ret.Problem = $true }
    if ($p.converted) { $ret.Converted++ }
    if ($p.inlinedBytes -gt 6MB) { $ret.Warnings += "Page $($p.n) pulls in $([Math]::Round($p.inlinedBytes / 1MB, 1)) MB of files; the presentation grows by about that much." }
  }
  foreach ($p in $Pages) {
    foreach ($x in @($p.renderLines)) { [void]$report.AppendLine($x) }
    if ($p.renderProblem) { $ret.Problem = $true }
  }
  if (@($Pages).Count -gt 1) {
    try {
      if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
      $pics = @($Pages | Where-Object { $_.png -and (Test-Path -LiteralPath $_.png) } | ForEach-Object { [string]$_.png })
      $n = $pics.Count; $cols = if ($n -le 4) { 2 } elseif ($n -le 9) { 3 } else { 4 }
      $ret.Sheet = [ImgKit.Compare]::Sheet([string[]]$pics, (Join-Path $pagesDir 'all_pages.png'), $cols, $(if ($cols -eq 2) { 960 } else { 640 }))
      if ($ret.Sheet) { [void]$report.AppendLine("all pages on one sheet (look at this first): $($ret.Sheet)") }
    } catch { }
  }
  $ret.Report = $report.ToString().TrimEnd()
  $ret.ReportFile = Join-Path $WorkDir 'page-report.txt'
  [IO.File]::WriteAllText($ret.ReportFile, $ret.Report, $script:Utf8NoBom)
  return $ret
}
