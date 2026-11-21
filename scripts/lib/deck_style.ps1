param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): reads the LOOK of a presentation that serves as a design reference.
# Theme colours and fonts, what the slides really use (backgrounds, text colours, fonts and sizes of titles and
# of other text, fills, lines, corner radius, shadows) and pictures of the first slides.
# The reference is opened read-only, without a window, with its macros disabled; it is never changed.
#   args: deck, outDir, maxSlides (pictures, default 6), scanSlides (statistics, default 40)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
function Hex([int]$ole) { return ('#{0:X2}{1:X2}{2:X2}' -f ($ole -band 255), (($ole -shr 8) -band 255), (($ole -shr 16) -band 255)) }
function Bump($map, [string]$key, [double]$w) { if (-not $key) { return }; if ($map.ContainsKey($key)) { $map[$key] += $w } else { $map[$key] = $w } }
function Ranked($map, [int]$n) {
  $out = @()
  foreach ($e in ($map.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $n)) { $out += [pscustomobject][ordered]@{ v = [string]$e.Key; w = [Math]::Round([double]$e.Value, 1) } }
  return $out
}
$pres = $null; $app = $null; $oldSec = $null
$st = @{ bg = @{}; text = @{}; fill = @{}; line = @{}; titleFont = @{}; bodyFont = @{}; titleSize = @{}; bodySize = @{}; radius = @{}
         sizeChars = @{}; sizeBold = @{}; sizeFont = @{}; textErr = ''
         boxes = 0; shadows = 0; pictures = 0; titleChars = 0; titleBold = 0; bodyChars = 0 }
$themeFonts = @{}
function FontName($font, [string]$sample) {
  # the name that applies to this text: East Asian for Hangul / CJK, Latin otherwise; theme placeholders are resolved
  $east = $false; foreach ($ch in $sample.ToCharArray()) { $cc = [int]$ch; if (($cc -ge 0x1100 -and $cc -le 0x11FF) -or ($cc -ge 0x3040 -and $cc -le 0x30FF) -or ($cc -ge 0x3130 -and $cc -le 0x318F) -or ($cc -ge 0x4E00 -and $cc -le 0x9FFF) -or ($cc -ge 0xAC00 -and $cc -le 0xD7A3)) { $east = $true; break } }
  $n = ''
  try { if ($east) { $n = [string]$font.NameFarEast } } catch {}
  if (-not $n) { try { $n = [string]$font.Name } catch {} }
  if ($n -match '^\+(mj|mn)-(lt|ea|cs)$') {
    $k = $(if ($Matches[1] -eq 'mj') { 'major' } else { 'minor' }) + $(if ($Matches[2] -eq 'ea') { 'East' } else { 'Latin' })
    if ($themeFonts[$k]) { $n = [string]$themeFonts[$k] }
  }
  return $n
}
function Scan($sh, [int]$depth) {
  $type = [int]$sh.Type
  if ($type -eq 6) {
    if ($depth -lt 2) { $gc = [Math]::Min([int]$sh.GroupItems.Count, 30); for ($g = 1; $g -le $gc; $g++) { try { Scan $sh.GroupItems.Item($g) ($depth + 1) } catch {} } }
    return
  }
  if ($type -eq 13 -or $type -eq 11) { $st.pictures++; return }
  $w = [double]$sh.Width; $h = [double]$sh.Height; $area = $w * $h
  $isTitle = $false
  if ($type -eq 14) { try { $pt = [int]$sh.PlaceholderFormat.Type; $isTitle = ($pt -eq 1 -or $pt -eq 3) } catch {} }
  try {
    if ($sh.Fill.Visible -ne 0 -and [int]$sh.Fill.Type -eq 1 -and [double]$sh.Fill.Transparency -lt 0.5 -and $area -gt 400) {
      Bump $st.fill (Hex ([int]$sh.Fill.ForeColor.RGB)) $area
      $st.boxes++
      try { if ($sh.Shadow.Visible -ne 0) { $st.shadows++ } } catch {}
      $ast = 0; try { $ast = [int]$sh.AutoShapeType } catch {}
      if ($ast -eq 5) { try { Bump $st.radius ([string][int][Math]::Round([double]$sh.Adjustments.Item(1) * [Math]::Min($w, $h) * 4 / 3)) 1 } catch {} }
      elseif ($ast -eq 1) { Bump $st.radius '0' 1 }
      elseif ($ast -eq 9) { Bump $st.radius 'round' 1 }
    }
  } catch {}
  try { if ($sh.Line.Visible -ne 0 -and [double]$sh.Line.Weight -gt 0) { Bump $st.line ((Hex ([int]$sh.Line.ForeColor.RGB)) + ' ' + [Math]::Round([double]$sh.Line.Weight, 1) + 'pt') 1 } } catch {}
  $has = $false; try { $has = ($sh.HasTextFrame -ne 0 -and $sh.TextFrame.HasText -ne 0) } catch {}
  if ($has) {
    # the text of a shape is sampled at its start and in its middle (walking every run of a large deck takes minutes)
    try {
      $tr = $sh.TextFrame.TextRange; $len = [int]$tr.Length; $all = [string]$tr.Text
      if ($len -ge 1 -and $all.Trim()) {
        $spots = @(1); if ($len -ge 8) { $spots += [int]($len / 2) }
        foreach ($pos in $spots) {
          $f = $tr.Characters($pos, 1).Font
          $from = [Math]::Max(0, $pos - 1)
          $sample = $all.Substring($from, [Math]::Min(20, $all.Length - $from))
          $name = FontName $f $sample
          $size = [int][Math]::Round([double]$f.Size); $color = Hex ([int]$f.Color.RGB); $part = $len / [double]$spots.Count
          Bump $st.text $color $part
          Bump $st.sizeChars ([string]$size) $part
          if ($f.Bold -ne 0) { Bump $st.sizeBold ([string]$size) $part }
          Bump $st.sizeFont ("$size|$name") $part
          if ($isTitle) { Bump $st.titleFont $name $part; Bump $st.titleSize ([string]$size) $part; $st.titleChars += $part; if ($f.Bold -ne 0) { $st.titleBold += $part } }
          else { Bump $st.bodyFont $name $part; Bump $st.bodySize ([string]$size) $part; $st.bodyChars += $part }
        }
      }
    } catch { if (-not $st.textErr) { $st.textErr = $_.Exception.Message; L "text of a shape could not be read: $($_.Exception.Message)" } }
  }
}
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}
  $pres = $app.Presentations.Open([string]$A.deck, -1, 0, 0)
  $SW = [double]$pres.PageSetup.SlideWidth; $SH = [double]$pres.PageSetup.SlideHeight
  $theme = [ordered]@{}
  $names = 'dark1', 'light1', 'dark2', 'light2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'link', 'visited'
  try { $tcs = $pres.SlideMaster.Theme.ThemeColorScheme; for ($i = 1; $i -le 12; $i++) { $theme[$names[$i - 1]] = Hex ([int]$tcs.Colors($i).RGB) } } catch { L "theme colours not read: $($_.Exception.Message)" }
  try {
    $tfs = $pres.SlideMaster.Theme.ThemeFontScheme
    $themeFonts['majorLatin'] = [string]$tfs.MajorFont.Item(1).Name; $themeFonts['majorEast'] = [string]$tfs.MajorFont.Item(2).Name
    $themeFonts['minorLatin'] = [string]$tfs.MinorFont.Item(1).Name; $themeFonts['minorEast'] = [string]$tfs.MinorFont.Item(2).Name
  } catch { L "theme fonts not read: $($_.Exception.Message)" }
  $scan = 40; if ($A.scanSlides) { $scan = [int]$A.scanSlides }
  $count = [Math]::Min([int]$pres.Slides.Count, $scan)
  for ($i = 1; $i -le $count; $i++) {
    $sl = $pres.Slides.Item($i)
    try {
      $f = $sl.Background.Fill; $t = [int]$f.Type
      if ($t -eq 3) { Bump $st.bg ('gradient ' + (Hex ([int]$f.ForeColor.RGB)) + ' ' + (Hex ([int]$f.BackColor.RGB))) 1 }
      elseif ($t -eq 4 -or $t -eq 6) { Bump $st.bg 'picture' 1 }
      else { Bump $st.bg (Hex ([int]$f.ForeColor.RGB)) 1 }
    } catch {}
    $n = [Math]::Min([int]$sl.Shapes.Count, 60)
    for ($k = 1; $k -le $n; $k++) { try { Scan $sl.Shapes.Item($k) 0 } catch {} }
  }
  $refs = @()
  $max = 6; if ($A.maxSlides) { $max = [int]$A.maxSlides }
  $pw = 1600; $ph = [int][Math]::Round($pw * $SH / $SW)
  for ($i = 1; $i -le [Math]::Min([int]$pres.Slides.Count, $max); $i++) {
    $p = Join-Path ([string]$A.outDir) ("ref_{0:D2}.png" -f $i)
    try { $pres.Slides.Item($i).Export($p, 'PNG', $pw, $ph); $refs += $p } catch { L "slide $i not exported: $($_.Exception.Message)" }
  }
  $shadowShare = 0.0; if ($st.boxes -gt 0) { $shadowShare = [Math]::Round($st.shadows / [double]$st.boxes, 2) }
  $titleBold = 0.0; if ($st.titleChars -gt 0) { $titleBold = [Math]::Round($st.titleBold / [double]$st.titleChars, 2) }
  # every text size that occurs: how much text, its usual font, how much of it is bold (largest first)
  $sizes = @()
  foreach ($e in ($st.sizeChars.GetEnumerator() | Sort-Object -Property @{ Expression = { [int]$_.Key } } -Descending)) {
    $pt = [string]$e.Key; $best = ''; $bw = 0.0
    foreach ($fe in $st.sizeFont.GetEnumerator()) { if (([string]$fe.Key).StartsWith("$pt|") -and [double]$fe.Value -gt $bw) { $bw = [double]$fe.Value; $best = ([string]$fe.Key).Substring($pt.Length + 1) } }
    $bold = 0.0; if ($st.sizeBold.ContainsKey($pt)) { $bold = [Math]::Round([double]$st.sizeBold[$pt] / [double]$e.Value, 2) }
    $sizes += [pscustomobject][ordered]@{ pt = [int]$pt; chars = [Math]::Round([double]$e.Value, 1); font = $best; boldShare = $bold }
  }
  L ("slides={0} scanned={1} pictures={2} boxes={3} sizes={4}" -f $pres.Slides.Count, $count, $refs.Count, $st.boxes, $sizes.Count)
  Write-JsonFile $A.result ([ordered]@{
    slideWidth = $SW; slideHeight = $SH; slides = [int]$pres.Slides.Count; scanned = $count
    theme = $theme; themeFonts = $themeFonts
    bg = @(Ranked $st.bg 6); text = @(Ranked $st.text 8); fill = @(Ranked $st.fill 10); line = @(Ranked $st.line 5)
    titleFont = @(Ranked $st.titleFont 4); bodyFont = @(Ranked $st.bodyFont 4); titleSize = @(Ranked $st.titleSize 5); bodySize = @(Ranked $st.bodySize 6)
    radius = @(Ranked $st.radius 5); shadowShare = $shadowShare; titleBoldShare = $titleBold; pictures = [int]$st.pictures
    titleChars = [int]$st.titleChars; bodyChars = [int]$st.bodyChars; sizes = @($sizes); refs = @($refs)
  })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
