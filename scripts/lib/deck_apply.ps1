param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): puts the stills of the pages on their slides, tags the slides, stores the
# pages inside the presentation (custom XML part) and saves it as a macro-enabled presentation (no macros yet).
#   args: deck (working copy; empty = new presentation), slideWidth, slideHeight, xmlFile, stageOut,
#         pages = [{ n, slide (0 = add a blank slide), shapeId (0 = none), left, top, width, height, png, bg, alt }]
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
$NS = 'urn:ppt-html:inline:v1'
$pres = $null; $app = $null; $oldSec = $null
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}
  if ($A.deck) {
    $pres = $app.Presentations.Open([string]$A.deck, 0, 0, 0)
  } else {
    $pres = $app.Presentations.Add(0)
    $pres.PageSetup.SlideWidth = [single]$A.slideWidth
    $pres.PageSetup.SlideHeight = [single]$A.slideHeight
  }
  # data of an earlier run
  try {
    $old = $pres.CustomXMLParts.SelectByNamespace($NS)
    for ($i = $old.Count; $i -ge 1; $i--) { $old.Item($i).Delete() }
  } catch { L "old data not removed: $($_.Exception.Message)" }
  for ($i = 1; $i -le $pres.Slides.Count; $i++) { try { $pres.Slides.Item($i).Tags.Delete('HtmlPage') } catch {} }

  $placed = @()
  foreach ($p in @($A.pages)) {
    $slide = $null
    if ([int]$p.slide -ge 1) { $slide = $pres.Slides.Item([int]$p.slide) }
    else {
      $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
      try {
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.Visible = -1
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = (ConvertTo-OleColor ([string]$p.bg))
      } catch { L "background of new slide not set: $($_.Exception.Message)" }
    }
    $z = 0
    # the shape that marked the place (or the still of an earlier run) gives way to the new still
    for ($k = $slide.Shapes.Count; $k -ge 1; $k--) {
      $sh = $slide.Shapes.Item($k)
      if (([int]$p.shapeId -gt 0 -and [int]$sh.Id -eq [int]$p.shapeId) -or $sh.Name -eq 'htmlArea') {
        if ([int]$sh.Id -eq [int]$p.shapeId) { $z = [int]$sh.ZOrderPosition }
        $sh.Delete()
      }
    }
    $pic = $slide.Shapes.AddPicture([string]$p.png, 0, -1, [single]$p.left, [single]$p.top, [single]$p.width, [single]$p.height)
    $pic.Name = 'htmlArea'
    $pic.AlternativeText = [string]$p.alt
    $pic.LockAspectRatio = 0
    $pic.Left = [single]$p.left; $pic.Top = [single]$p.top; $pic.Width = [single]$p.width; $pic.Height = [single]$p.height
    if ($z -gt 0) { $guard = 0; while ([int]$pic.ZOrderPosition -gt $z -and $guard -lt 500) { $pic.ZOrder(3); $guard++ } }
    [void]$slide.Tags.Add('HtmlPage', [string]$p.n)
    # other shapes under the area: during the show the live page is always on top of them
    $under = @()
    $aL = [double]$pic.Left; $aT = [double]$pic.Top; $aR = $aL + [double]$pic.Width; $aB = $aT + [double]$pic.Height
    for ($k = 1; $k -le $slide.Shapes.Count; $k++) {
      $sh = $slide.Shapes.Item($k)
      try {
        if ($sh.Name -eq 'htmlArea' -or $sh.Visible -eq 0) { continue }
        $txt = ''
        if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = ([string]$sh.TextFrame.TextRange.Text -replace '\s+', ' ').Trim() }
        if ([int]$sh.Type -eq 14 -and -not $txt) { continue }
        $sL = [double]$sh.Left; $sT = [double]$sh.Top; $sR = $sL + [double]$sh.Width; $sB = $sT + [double]$sh.Height
        $ix = [Math]::Min($aR, $sR) - [Math]::Max($aL, $sL); $iy = [Math]::Min($aB, $sB) - [Math]::Max($aT, $sT)
        if ($ix -le 2 -or $iy -le 2) { continue }
        $whole = ($sL -ge $aL - 1 -and $sT -ge $aT - 1 -and $sR -le $aR + 1 -and $sB -le $aB + 1)
        $label = [string]$sh.Name
        if ($txt) { if ($txt.Length -gt 30) { $txt = $txt.Substring(0, 30) + '...' }; $label += " `"$txt`"" }
        $under += ($label + $(if ($whole) { ' (hidden)' } else { ' (partly covered)' }))
      } catch {}
    }
    $placed += [pscustomobject][ordered]@{ n = $p.n; slide = [int]$slide.SlideIndex; left = [double]$pic.Left; top = [double]$pic.Top; width = [double]$pic.Width; height = [double]$pic.Height; covers = @($under) }
    L ("page {0} -> slide {1} at {2:N1},{3:N1} {4:N1}x{5:N1} pt" -f $p.n, $slide.SlideIndex, $pic.Left, $pic.Top, $pic.Width, $pic.Height)
  }
  [void]$pres.Tags.Add('InlineHtml', '1')
  $xml = [IO.File]::ReadAllText([string]$A.xmlFile, [Text.Encoding]::UTF8)
  [void]$pres.CustomXMLParts.Add($xml)
  if (Test-Path -LiteralPath ([string]$A.stageOut)) { Remove-Item -LiteralPath ([string]$A.stageOut) -Force }
  $pres.SaveAs([string]$A.stageOut, 25)
  L "saved $($A.stageOut) slides=$($pres.Slides.Count)"
  Write-JsonFile $A.result ([ordered]@{ slideCount = [int]$pres.Slides.Count; placed = @($placed) })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
