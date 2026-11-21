# Dot-source (after common.ps1 and pages.ps1). Prepares HTML pages for the modern engine of the slide show
# (Microsoft Edge WebView2): one self-contained document per page, exactly as the page was written (nothing is
# converted, nothing is checked against the rules of the old engine), fonts embedded, the picture for the slide
# taken in Edge, which is the same engine. Plain ASCII on purpose (see hd_core.ps1).

# Pages: objects with  html (source path), n (1..), canvasW, canvasH (CSS px), label.
# Adds to each page: docM (document), pngM (picture, Scale x the canvas), bgM (#RRGGBB), titleM, modernLines, modernProblem, scrolls.
# Returns: Ok ($false when Edge cannot be used as the build tool: Why says why), Problem, Report, Fatal, FatalCode, Sheet.
function Invoke-ModernPrep {
  param([Parameter(Mandatory)]$Pages, [Parameter(Mandatory)][string]$WorkDir, [int]$SettleMs = 1500, [int]$Scale = 2, [string]$Fonts = 'auto', [string]$Images = 'auto', [string]$Probe = '')
  $ret = [pscustomobject]@{ Ok = $false; Why = ''; Problem = $false; Report = ''; Fatal = $null; FatalCode = 0; Sheet = ''; ReportFile = '' }
  Initialize-HtmlPrep
  $headH = Read-Asset 'harness-head-wv.txt'; $tailH = Read-Asset 'harness-tail-wv.txt'
  $pagesDir = Join-Path $WorkDir 'pages'
  if (-not (Test-Path -LiteralPath $pagesDir)) { New-Item -ItemType Directory -Force $pagesDir | Out-Null }
  $denied = "ACCESS-DENIED: PowerShell is not allowed to read '{0}' on this PC (security software). Copy the page and the files it uses to a folder under '$(Join-Path (Get-HdRoot) 'in')' with your own file tools, then run again with that copy."
  $wantFonts = (@('off', 'never', 'false') -notcontains $Fonts.ToLowerInvariant())
  $edge = Start-EdgeTool $WorkDir
  if (-not $edge) { $ret.Why = Get-EdgeWhyNot; return $ret }
  $ret.Ok = $true
  $again = '    note: the tool browser (Edge) stopped answering in the middle of the run and was started again'
  try {
    foreach ($p in $Pages) {
      $out = New-Object 'Collections.Generic.List[string]'
      $src = $null; $st = $null
      try { $src = [PptHtml.HtmlPrep]::ReadPage($p.html) } catch { $ret.Fatal = ($denied -f $p.html); $ret.FatalCode = 3; return $ret }
      [PptHtml.HtmlPrep]::ImageMaxSide = $(if (@('keep', 'off', 'never', 'false') -contains $Images.ToLowerInvariant()) { 0 } else { [Math]::Max(1920, 3 * [Math]::Max([int]$p.canvasW, [int]$p.canvasH)) })
      try { $st = [PptHtml.HtmlPrep]::Begin($src, (Split-Path -Parent $p.html)) }
      catch {
        $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
        if ($e -is [UnauthorizedAccessException]) { $ret.Fatal = ($denied -f "a file used by $($p.html): $($e.Message)"); $ret.FatalCode = 3 } else { $ret.Fatal = "Preparing $($p.html) failed: $($e.Message)"; $ret.FatalCode = 1 }
        return $ret
      }
      $png = Join-Path $pagesDir ("page_{0:D2}.png" -f $p.n)
      $file = Join-Path $pagesDir ("page_{0:D2}.html" -f $p.n)
      $docM = ''; $info = $null; $consoleErr = @(); $probeText = ''; $lines = @()
      for ($try = 0; $try -lt 2; $try++) {
        try {
          if (-not (Test-EdgeTool)) { $edge = Restart-EdgeTool $WorkDir; if (-not $edge) { throw 'the tool browser could not be started again' }; $lines += $again }
          $edge.SetViewport([int]$p.canvasW, [int]$p.canvasH, [double]$Scale)
          if ($wantFonts) {
            [void](Invoke-EdgeTool $edge 'IHTool.show' @(("<!DOCTYPE html>`n" + $st.Doc), [int]$p.canvasW, [int]$p.canvasH, 200) 60000)
            foreach ($x in @(Add-PageFonts $edge $st)) { $st.R.Notes.Add([string]$x) }
          }
          $docM = [PptHtml.HtmlPrep]::FinishModern($st, $headH, $tailH)
          $before = $edge.Errors.Count
          [void](Invoke-EdgeTool $edge 'IHTool.show' @($docM, [int]$p.canvasW, [int]$p.canvasH, $SettleMs) 90000)
          $edge.Screenshot($png)
          $info = ConvertFrom-BigJson ([string]$edge.Eval('IHTool.info()', $false, 15000))
          if ($Probe) { try { $probeText = [string]$edge.Eval("(function(){try{var w=document.querySelector('iframe').contentWindow;return String(w.eval($(ConvertTo-BigJson $Probe)));}catch(e){return 'PROBE-ERR '+e.message;}})()", $false, 15000) } catch { $probeText = "PROBE-ERR $($_.Exception.Message)" } }
          for ($k = $before; $k -lt $edge.Errors.Count -and $k -lt $before + 4; $k++) { $consoleErr += [string]$edge.Errors[$k] }
          [void]$edge.Eval('IHTool.hide()', $false, 10000)
          break
        } catch {
          $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
          if ($try -eq 0 -and -not (Test-EdgeTool)) { continue }
          $ret.Fatal = "Preparing $($p.html) for the modern engine failed: $($e.Message)"; $ret.FatalCode = 1; return $ret
        }
      }
      [IO.File]::WriteAllText($file, $docM, $script:Utf8NoBom)
      $bg = ''; if ($info -and $info['bg']) { $bg = [string]$info['bg'] }
      if (-not $bg) {
        # no plain background colour (a gradient, a picture ...): the colour the picture mostly has
        try {
          if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
          $pal = [ImgKit.Compare]::Palette($png, 1); if ($pal -match '^([0-9A-F]{6}):') { $bg = '#' + $Matches[1] }
        } catch {}
      }
      if (-not $bg) { $bg = '#FFFFFF' }
      $problem = $false
      $out.Add(("page {0}: {1}  ->  {2}, canvas {3} x {4} px  [modern engine: used as written]" -f $p.n, $p.html, $p.label, $p.canvasW, $p.canvasH))
      foreach ($x in $st.R.Notes) { $out.Add("    note: $x") }
      foreach ($x in $lines) { $out.Add($x) }
      foreach ($x in $st.R.Missing) { $out.Add("    ERROR missing file: $x"); $problem = $true }
      foreach ($x in $st.R.External) { $out.Add("    warn  needs internet during the show: $x") }
      $errText = ''; if ($info) { $errText = [string]$info['errors'] }
      if ($errText) { $out.Add("    ERROR script error(s) while loading: $errText"); $problem = $true }
      elseif ($info -and -not [bool]$info['ready']) { $out.Add('    ERROR the page script did not run to the end'); $problem = $true }
      foreach ($x in $consoleErr) { if (-not $errText -or $errText.IndexOf($x.Substring(0, [Math]::Min(30, $x.Length))) -lt 0) { $out.Add("    warn  the browser console reports: $x") } }
      $scrolls = $false
      if ($info -and [int]$info['clientH'] -gt 0) {
        if ([int]$info['scrollH'] -gt [int]$info['clientH'] + 2) { $scrolls = $true; $out.Add("    note: the page is $($info['scrollH']) px tall, the slide shows $($info['clientH']) px of it at a time: it scrolls inside the slide (mouse wheel, touch, keys). The picture shows the top.") }
        if ([int]$info['scrollW'] -gt [int]$info['clientW'] + 2) { $out.Add("    warn  the page is $($info['scrollW']) px wide but the slide area is $($info['clientW']) px: it scrolls sideways") }
      }
      if ($Probe) { $out.Add("    probe: $probeText") }
      $out.Add("    picture (${Scale}x the canvas, drawn by Edge, the engine of the slide show): $png   background $bg")
      $p | Add-Member -NotePropertyName docM -NotePropertyValue $docM -Force
      $p | Add-Member -NotePropertyName pngM -NotePropertyValue $png -Force
      $p | Add-Member -NotePropertyName fileM -NotePropertyValue $file -Force
      $p | Add-Member -NotePropertyName bgM -NotePropertyValue $bg -Force
      $p | Add-Member -NotePropertyName titleM -NotePropertyValue $(if ($info) { [string]$info['title'] } else { '' }) -Force
      $p | Add-Member -NotePropertyName scrolls -NotePropertyValue $scrolls -Force
      $p | Add-Member -NotePropertyName modernLines -NotePropertyValue $out -Force
      $p | Add-Member -NotePropertyName modernProblem -NotePropertyValue $problem -Force
      $p | Add-Member -NotePropertyName inlinedBytesM -NotePropertyValue ([long]$st.R.InlinedBytes) -Force
      if ($problem) { $ret.Problem = $true }
    }
  } finally { Stop-EdgeTool }
  $report = New-Object Text.StringBuilder
  foreach ($p in $Pages) { foreach ($x in $p.modernLines) { [void]$report.AppendLine($x) } }
  if (@($Pages).Count -gt 1) {
    try {
      if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
      $pics = @($Pages | Where-Object { $_.pngM -and (Test-Path -LiteralPath $_.pngM) } | ForEach-Object { [string]$_.pngM })
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
