param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): reads a presentation for the web version. Every slide is exported as a
# picture; titles, text, speaker notes, hyperlinks and (if the presentation was built by this skill) the live
# HTML pages and their places are collected.
#   args: file, outDir, width (px), imageFormat (png | jpg)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
$NS = 'urn:ppt-html:inline:v1'
$pres = $null; $app = $null; $oldSec = $null
function Clean([string]$s) { if (-not $s) { return '' }; return (($s -replace "[\x0B\r]", "`n") -replace "[\x00-\x08\x0C\x0E-\x1F]", '').Trim() }
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}
  $pres = $app.Presentations.Open([string]$A.file, -1, 0, 0)
  $sw = [double]$pres.PageSetup.SlideWidth; $sh = [double]$pres.PageSetup.SlideHeight
  $w = 1600; if ($A.width) { $w = [int]$A.width }
  $h = [int][Math]::Round($w * $sh / $sw)
  $fmt = 'PNG'; $ext = 'png'
  if ("$($A.imageFormat)" -eq 'jpg') { $fmt = 'JPG'; $ext = 'jpg' }
  $slides = @()
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $img = "slide_{0:D3}.$ext" -f $i
    $sl.Export((Join-Path ([string]$A.outDir) $img), $fmt, $w, $h)
    $title = ''
    try { if ($sl.Shapes.HasTitle) { $title = Clean ([string]$sl.Shapes.Title.TextFrame.TextRange.Text) } } catch {}
    $texts = @()
    for ($k = 1; $k -le $sl.Shapes.Count; $k++) {
      $shp = $sl.Shapes.Item($k)
      try {
        if ($shp.HasTextFrame -and $shp.TextFrame.HasText) { $texts += (Clean ([string]$shp.TextFrame.TextRange.Text)) }
        elseif ($shp.HasTable) {
          $tb = $shp.Table
          for ($r = 1; $r -le $tb.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $tb.Columns.Count; $c++) { $row += (Clean ([string]$tb.Cell($r, $c).Shape.TextFrame.TextRange.Text)) }
            $texts += ($row -join ' | ')
          }
        }
      } catch {}
    }
    $notes = ''
    try { $notes = Clean ([string]$sl.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text) } catch {}
    $links = @()
    try {
      for ($k = 1; $k -le $sl.Hyperlinks.Count; $k++) {
        $hl = $sl.Hyperlinks.Item($k)
        $addr = [string]$hl.Address; $sub = [string]$hl.SubAddress
        $target = 0
        if (-not $addr -and $sub -match '^\d+,(\d+),') { $target = [int]$Matches[1] }
        if (-not $addr -and -not $target) { continue }
        $box = $null
        try {
          $owner = $hl.Parent.Parent
          if ([int]$hl.Type -eq 1) { $box = @([double]$owner.Left, [double]$owner.Top, [double]$owner.Width, [double]$owner.Height) }
          else { $box = @([double]$owner.BoundLeft, [double]$owner.BoundTop, [double]$owner.BoundWidth, [double]$owner.BoundHeight) }
        } catch {}
        if ($box -and $box[2] -gt 1 -and $box[3] -gt 1) { $links += [pscustomobject][ordered]@{ left = $box[0]; top = $box[1]; width = $box[2]; height = $box[3]; href = $addr; slide = $target } }
      }
    } catch {}
    $page = 0; $area = $null; $state = ''
    try {
      $tag = [string]$sl.Tags.Item('HtmlPage')
      if ($tag) {
        $ha = $sl.Shapes.Item('htmlArea')
        $page = [int]$tag
        $area = [pscustomobject][ordered]@{ left = [double]$ha.Left; top = [double]$ha.Top; width = [double]$ha.Width; height = [double]$ha.Height }
      }
    } catch { $page = 0; $area = $null }
    # what the page stored during a show (edited text, votes): pieces in the slide's tags
    if ($page -gt 0) {
      try {
        $cnt = 0; [void][int]::TryParse([string]$sl.Tags.Item('HtmlState0'), [ref]$cnt)
        if ($cnt -gt 0 -and $cnt -le 200) { $sb = New-Object Text.StringBuilder; for ($q = 1; $q -le $cnt; $q++) { [void]$sb.Append([string]$sl.Tags.Item('HtmlState' + $q)) }; $state = $sb.ToString() }
      } catch { $state = '' }
    }
    $hidden = $false; try { $hidden = ($sl.SlideShowTransition.Hidden -ne 0) } catch {}
    $slides += [pscustomobject][ordered]@{ index = $i; title = $title; hidden = $hidden; image = $img; text = ($texts -join "`n"); notes = $notes; links = @($links); page = $page; area = $area; state = $state }
  }
  $pagesXml = ''
  try {
    $parts = $pres.CustomXMLParts.SelectByNamespace($NS)
    if ($parts.Count -ge 1) {
      $pagesXml = Join-Path ([string]$A.outDir) 'pages.xml'
      [IO.File]::WriteAllText($pagesXml, [string]$parts.Item(1).XML, (New-Object Text.UTF8Encoding($false)))
    }
  } catch { L "page data not read: $($_.Exception.Message)" }
  L "exported $($slides.Count) slide(s) at ${w}x$h ($fmt); live pages: $(@($slides | Where-Object { $_.page -gt 0 }).Count)"
  Write-JsonFile $A.result ([ordered]@{ name = [string]$pres.Name; slideWidth = $sw; slideHeight = $sh; imageWidth = $w; imageHeight = $h; slides = @($slides); pagesXml = $pagesXml })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
