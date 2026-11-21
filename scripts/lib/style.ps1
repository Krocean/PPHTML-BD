# Dot-source (after common.ps1, pages.ps1 and fonts.ps1). The look of a deck as a handful of values:
# colour arithmetic, the full set of theme variables of the design kit derived from a few base values,
# the theme file, the design description (design.md) and a preview picture.
# Plain ASCII on purpose (see hd_core.ps1).

# Hangul font names, written as character codes because these scripts are plain ASCII.
function Get-CodeText([int[]]$Codes) { return (-join ($Codes | ForEach-Object { [char]$_ })) }
# Malgun Gothic, Batang, Gulim, Dotum, Gungsuh under their Korean names: fonts every Windows PC has
$script:KoEverywhere = @((Get-CodeText @(0xB9D1, 0xC740, 0x20, 0xACE0, 0xB515)), (Get-CodeText @(0xBC14, 0xD0D5)), (Get-CodeText @(0xAD74, 0xB9BC)), (Get-CodeText @(0xB3CB, 0xC6C0)), (Get-CodeText @(0xAD81, 0xC11C)))
# the words for a serif face in Korean font names: myeongjo, batang
$script:KoSerif = @((Get-CodeText @(0xBA85, 0xC870)), (Get-CodeText @(0xBC14, 0xD0D5)))

function ConvertTo-Rgb([string]$Color) {
  $c = ([string]$Color).Trim()
  if ($c -match '^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$') { return @([Convert]::ToInt32($Matches[1], 16), [Convert]::ToInt32($Matches[2], 16), [Convert]::ToInt32($Matches[3], 16)) }
  if ($c -match '^#?([0-9a-fA-F])([0-9a-fA-F])([0-9a-fA-F])$') { return @(([Convert]::ToInt32($Matches[1], 16) * 17), ([Convert]::ToInt32($Matches[2], 16) * 17), ([Convert]::ToInt32($Matches[3], 16) * 17)) }
  if ($c -match '^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)') { return @([int][Math]::Round([double]$Matches[1]), [int][Math]::Round([double]$Matches[2]), [int][Math]::Round([double]$Matches[3])) }
  return $null
}
function Test-Color([string]$Color) { return ($null -ne (ConvertTo-Rgb $Color)) }
function Format-Hex($Rgb) {
  $v = @(); foreach ($x in @($Rgb)[0..2]) { $v += [int][Math]::Max(0, [Math]::Min(255, [Math]::Round([double]$x))) }
  return ('#{0:X2}{1:X2}{2:X2}' -f $v[0], $v[1], $v[2])
}
function Get-Mix([string]$From, [string]$To, [double]$T) {
  $p = ConvertTo-Rgb $From; $q = ConvertTo-Rgb $To
  return (Format-Hex @(($p[0] + ($q[0] - $p[0]) * $T), ($p[1] + ($q[1] - $p[1]) * $T), ($p[2] + ($q[2] - $p[2]) * $T)))
}
function Get-Luminance([string]$Color) {
  $p = ConvertTo-Rgb $Color; $l = @()
  foreach ($v in $p) { $s = $v / 255.0; if ($s -le 0.03928) { $l += ($s / 12.92) } else { $l += [Math]::Pow(($s + 0.055) / 1.055, 2.4) } }
  return (0.2126 * $l[0] + 0.7152 * $l[1] + 0.0722 * $l[2])
}
function Get-Contrast([string]$One, [string]$Two) {
  $x = Get-Luminance $One; $y = Get-Luminance $Two
  if ($x -lt $y) { $t = $x; $x = $y; $y = $t }
  return (($x + 0.05) / ($y + 0.05))
}
function Get-Saturation([string]$Color) {
  $p = ConvertTo-Rgb $Color; $mx = [Math]::Max($p[0], [Math]::Max($p[1], $p[2])); $mn = [Math]::Min($p[0], [Math]::Min($p[1], $p[2]))
  if ($mx -le 0) { return 0.0 }
  return (($mx - $mn) / [double]$mx)
}
function Get-Hue([string]$Color) {
  $p = ConvertTo-Rgb $Color; $r = $p[0] / 255.0; $g = $p[1] / 255.0; $b = $p[2] / 255.0
  $mx = [Math]::Max($r, [Math]::Max($g, $b)); $mn = [Math]::Min($r, [Math]::Min($g, $b)); $d = $mx - $mn
  if ($d -lt 0.0001) { return 0.0 }
  if ($mx -eq $r) { $h = (($g - $b) / $d) % 6 } elseif ($mx -eq $g) { $h = (($b - $r) / $d) + 2 } else { $h = (($r - $g) / $d) + 4 }
  $h = $h * 60; if ($h -lt 0) { $h += 360 }
  return $h
}
function Get-Rgba([string]$Color, [double]$Alpha) {
  $p = ConvertTo-Rgb $Color
  return ('rgba({0},{1},{2},{3})' -f $p[0], $p[1], $p[2], $Alpha.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture))
}
# True when two colours would be hard to tell apart in a chart.
function Test-ColorNear([string]$One, [string]$Two) {
  $p = ConvertTo-Rgb $One; $q = ConvertTo-Rgb $Two
  $dist = [Math]::Sqrt([Math]::Pow($p[0] - $q[0], 2) + [Math]::Pow($p[1] - $q[1], 2) + [Math]::Pow($p[2] - $q[2], 2))
  if ($dist -lt 48) { return $true }
  if ((Get-Saturation $One) -gt 0.3 -and (Get-Saturation $Two) -gt 0.3) {
    $dh = [Math]::Abs((Get-Hue $One) - (Get-Hue $Two)); if ($dh -gt 180) { $dh = 360 - $dh }
    if ($dh -lt 16 -and $dist -lt 110) { return $true }
  }
  return $false
}

# A font stack the build can embed: the wanted families that exist on this PC or in the skill's font pack, then
# a family with Hangul, then the generic one. Notes say what was replaced.
function Get-StyleFontStack([string[]]$Names, [string]$Kind = 'sans') {
  $idx = Get-FontIndex
  $kept = @(); $notes = @(); $hangul = $false; $serif = ($Kind -eq 'serif')
  foreach ($raw in @($Names)) {
    $n = ([string]$raw).Trim().Trim('"').Trim("'").Trim()
    if (-not $n -or $kept -contains $n) { continue }
    if ($n -match '^(sans-serif|serif|monospace|cursive|system-ui|ui-sans-serif|ui-serif|ui-monospace|-apple-system|BlinkMacSystemFont|inherit|initial|unset|emoji)$') { continue }
    if ($n -match '^\+m[jn]-') { continue }
    if (($n -match 'serif|myeongjo|batang|times|georgia|garamond|playfair|didot|bodoni|mincho' -or $n.Contains($script:KoSerif[0]) -or $n.Contains($script:KoSerif[1])) -and $n -notmatch 'sans') { $serif = $true }
    $key = $n.ToLowerInvariant()
    # fonts that every Windows PC has are kept even when they may not be embedded (PowerPoint finds them on the PC)
    $everywhere = ($key -match '^(georgia|segoe ui|segoe ui semibold|malgun gothic|arial|times new roman|verdana|tahoma|trebuchet ms|calibri|cambria|consolas|batang|gulim|dotum|gungsuh)$') -or ($script:KoEverywhere -contains $n)
    if ($idx.ContainsKey($key)) {
      $faces = @($idx[$key]); $ok = @($faces | Where-Object { $_.Embeddable })
      if ($ok.Count -gt 0 -or $everywhere) {
        $kept += $n
        try { if (Test-FaceHangul $(if ($ok.Count -gt 0) { $ok[0] } else { $faces[0] })) { $hangul = $true } } catch {}
        if ($ok.Count -eq 0 -and $kept.Count -eq 1) { $notes += "'$n' is a Windows font that is not embedded: PowerPoint on Windows shows it, a browser on another system falls back to the next font" }
        if ($kept.Count -ge 2) { break }
      } elseif ($kept.Count -eq 0) { $notes += "'$n': on this PC, but its licence flag does not allow embedding - replaced" }
    } elseif ($kept.Count -eq 0) { $notes += "'$n': not on this PC and not in the font pack - replaced" }
  }
  $fallback = 'Pretendard'; if ($serif) { $fallback = 'Noto Serif KR' }
  if ((-not $hangul -or $kept.Count -eq 0) -and ($kept -notcontains $fallback)) { $kept += $fallback }
  $generic = 'sans-serif'; if ($serif) { $generic = 'serif' }
  $stack = (($kept | ForEach-Object { '"' + $_ + '"' }) -join ', ') + ', ' + $generic
  return [pscustomobject]@{ stack = $stack; notes = @($notes); serif = $serif; first = [string]$kept[0] }
}

# The whole theme from a few base values. Keys of $S (all optional):
#   bg surface ink accent  colours (#RRGGBB)        colors  more chart colours, in order
#   display body           font stacks              displayWeight strongWeight  100-900
#   tracking               headline letter-spacing in em (-0.02)
#   radius                 corner of cards in px    buttonRadius  px, or 'pill'
#   shadow                 soft | lift | border | none        button  accent | ink
#   labelCase              none | uppercase         good bad warn  status colours
# Returns: tokens (ordered name -> value), notes, dark.
function New-ThemeTokens([hashtable]$S) {
  $notes = @(); $white = '#FFFFFF'; $black = '#000000'
  $bg = '#FFFFFF'; if ($S.bg -and (Test-Color $S.bg)) { $bg = Format-Hex (ConvertTo-Rgb $S.bg) }
  $dark = ((Get-Luminance $bg) -lt 0.2)
  $ink = '#111827'; if ($dark) { $ink = '#F2F4F7' }
  if ($S.ink -and (Test-Color $S.ink)) {
    $cand = Format-Hex (ConvertTo-Rgb $S.ink)
    if ((Get-Contrast $cand $bg) -ge 4.5) { $ink = $cand } else { $notes += "text colour $cand has too little contrast on $bg (" + ([Math]::Round((Get-Contrast $cand $bg), 1)) + ":1): $ink is used instead" }
  }
  $accent = '#2563EB'; if ($dark) { $accent = '#7C8CFF' }
  if ($S.accent -and (Test-Color $S.accent)) { $accent = Format-Hex (ConvertTo-Rgb $S.accent) }
  if ((Get-Contrast $accent $bg) -lt 1.6) { $notes += "accent $accent is hard to see on $bg (" + ([Math]::Round((Get-Contrast $accent $bg), 1)) + ":1)" }
  if ($S.surface -and (Test-Color $S.surface)) { $surface = Format-Hex (ConvertTo-Rgb $S.surface) }
  elseif ($dark) { $surface = Get-Mix $bg $white 0.06 }
  else { $surface = '#FFFFFF' }
  $bg2 = Get-Mix $bg $ink 0.045; if ($dark) { $bg2 = Get-Mix $bg $white 0.035 }
  $ink2 = Get-Mix $ink $bg 0.34; $ink3 = Get-Mix $ink $bg 0.56
  $lineT = 0.87; if ($dark) { $lineT = 0.82 }
  $line = Get-Mix $ink $bg $lineT
  $accentInk = '#111111'; if ((Get-Contrast $accent $white) -ge (Get-Contrast $accent '#111111')) { $accentInk = '#FFFFFF' }
  $softT = 0.88; if ($dark) { $softT = 0.80 }
  $accentSoft = Get-Mix $accent $bg $softT
  $deep = $accent; $step = 0
  while ((Get-Contrast $deep $bg) -lt 4.5 -and $step -lt 14) { $deep = Get-Mix $deep $ink 0.12; $step++ }
  if (-not $dark -and $step -eq 0) { $deep = Get-Mix $accent $black 0.14 }
  $status = @{ good = '#15803D'; bad = '#DC2626'; warn = '#B45309' }
  if ($dark) { $status = @{ good = '#4ADE80'; bad = '#F87171'; warn = '#FBBF24' } }
  foreach ($k in 'good', 'bad', 'warn') { if ($S[$k] -and (Test-Color $S[$k])) { $status[$k] = Format-Hex (ConvertTo-Rgb $S[$k]) } }
  $chart = @($accent)
  $spare = @('#0EA5A4', '#F59E0B', '#8B5CF6', '#EF4444', '#64748B', '#22C55E', '#EC4899')
  if ($dark) { $spare = @('#2DD4BF', '#FBBF24', '#A78BFA', '#F87171', '#94A3B8', '#4ADE80', '#F472B6') }
  foreach ($c in (@($S.colors) + $spare)) {
    if ($chart.Count -ge 6) { break }
    if (-not $c -or -not (Test-Color $c)) { continue }
    $h = Format-Hex (ConvertTo-Rgb $c)
    if ((Get-Contrast $h $surface) -lt 1.35) { continue }
    $near = $false; foreach ($e in $chart) { if (Test-ColorNear $h $e) { $near = $true; break } }
    if (-not $near) { $chart += $h }
  }
  $press = $accent; $pressInk = $accentInk
  if ([string]$S.button -eq 'ink') { $press = $ink; $pressInk = $bg }
  $display = '"Pretendard", sans-serif'; if ($S.display) { $display = [string]$S.display }
  $body = '"Pretendard", sans-serif'; if ($S.body) { $body = [string]$S.body }
  $dw = 700; if ($S.displayWeight) { $dw = [int]$S.displayWeight }
  $dw = [int]([Math]::Max(100, [Math]::Min(900, [Math]::Round($dw / 100.0) * 100)))
  $sw = 600; if ($S.strongWeight) { $sw = [int]$S.strongWeight }
  $sw = [int]([Math]::Max(400, [Math]::Min(900, [Math]::Round($sw / 100.0) * 100)))
  $tr = -0.02; if ($dw -lt 500) { $tr = -0.01 }
  if ($null -ne $S.tracking -and "$($S.tracking)" -ne '') { $tr = [Math]::Max(-0.06, [Math]::Min(0.1, [double]$S.tracking)) }
  $trText = '0'; if ([Math]::Abs($tr) -ge 0.0005) { $trText = $tr.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture) + 'em' }
  $upper = ([string]$S.labelCase -eq 'uppercase')
  $radius = 14; if ($null -ne $S.radius -and "$($S.radius)" -ne '') { $radius = [int][Math]::Max(0, [Math]::Min(36, [double]$S.radius)) }
  $radiusS = [int][Math]::Round($radius * 0.6)
  $radiusBtn = "${radiusS}px"
  if ("$($S.buttonRadius)" -eq 'pill') { $radiusBtn = '999px' }
  elseif ($null -ne $S.buttonRadius -and "$($S.buttonRadius)" -ne '') { $radiusBtn = ([int][Math]::Max(0, [Math]::Min(60, [double]$S.buttonRadius))).ToString() + 'px' }
  $shadowKind = [string]$S.shadow
  if (-not $shadowKind) { $shadowKind = 'soft'; if ($dark) { $shadowKind = 'border' } }
  $shadeInk = $ink; if ($dark) { $shadeInk = $black }
  $cardBorder = '0 solid transparent'
  switch ($shadowKind) {
    'none' { $shadow = '0 0 0 0 transparent'; if ($surface -eq $bg) { $cardBorder = "1px solid $line" } }
    'border' { $shadow = '0 0 0 0 transparent'; $cardBorder = "1px solid $line" }
    'lift' { $shadow = "0 1px 2px $(Get-Rgba $shadeInk 0.10), 0 6px 16px -6px $(Get-Rgba $shadeInk 0.22)" }
    default { $shadow = "0 0 0 1px $(Get-Rgba $ink 0.06), 0 2px 6px -2px $(Get-Rgba $shadeInk 0.14)" }
  }
  $popA = 0.28; if ($dark) { $popA = 0.7 }
  $shadowPop = "0 0 0 1px $(Get-Rgba $ink 0.08), 0 12px 28px -8px $(Get-Rgba $shadeInk $popA)"
  $hoverA = 0.05; $hover2A = 0.09; if ($dark) { $hoverA = 0.07; $hover2A = 0.12 }
  $labelTr = '0.04em'; $labelCase = 'none'; if ($upper) { $labelTr = '0.08em'; $labelCase = 'uppercase' }
  $t = [ordered]@{}
  $t['bg'] = $bg; $t['bg2'] = $bg2; $t['surface'] = $surface; $t['ink'] = $ink; $t['ink2'] = $ink2; $t['ink3'] = $ink3; $t['line'] = $line
  $t['accent'] = $accent; $t['accent-ink'] = $accentInk; $t['accent-soft'] = $accentSoft; $t['accent-deep'] = $deep
  foreach ($k in 'good', 'bad', 'warn') { $t[$k] = $status[$k]; $t["$k-soft"] = Get-Mix $status[$k] $bg $softT }
  for ($i = 0; $i -lt 6; $i++) { $t["c$($i + 1)"] = $chart[[Math]::Min($i, $chart.Count - 1)] }
  $t['hover'] = Get-Rgba $ink $hoverA; $t['hover2'] = Get-Rgba $ink $hover2A
  $t['press'] = $press; $t['press-ink'] = $pressInk
  $t['display'] = $display; $t['display-weight'] = [string]$dw; $t['display-tracking'] = $trText
  $t['body'] = $body; $t['strong-weight'] = [string]$sw; $t['label-tracking'] = $labelTr; $t['label-case'] = $labelCase
  $t['radius'] = "${radius}px"; $t['radius-s'] = "${radiusS}px"; $t['radius-btn'] = $radiusBtn
  $t['card-border'] = $cardBorder; $t['shadow'] = $shadow; $t['shadow-pop'] = $shadowPop
  if ((Get-Contrast $ink2 $bg) -lt 4.0) { $notes += 'secondary text is light on this background: keep it for short labels only' }
  return [pscustomobject]@{ tokens = $t; notes = @($notes); dark = $dark }
}

function Write-ThemeCss($Tokens, [string]$Name, [string]$Note, [string]$Path) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.AppendLine("/* ppt-html design kit - theme `"$Name`": $Note")
  [void]$sb.AppendLine('   Made by new-theme.ps1 / get-style.ps1 from a few base values; change any value by hand.')
  [void]$sb.AppendLine('   Link ONE theme file before kit.css (keep every name; each is defined exactly once, here). */')
  [void]$sb.AppendLine(':root {')
  foreach ($k in $Tokens.Keys) { [void]$sb.AppendLine("  --ph-${k}: $($Tokens[$k]);") }
  [void]$sb.AppendLine('}')
  [IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
}

# Draws the specimen page (every part of the kit on one page) with a theme file. Returns the picture path or ''.
function New-StylePreview([string]$ThemeFile, [string]$OutPng, [string]$WorkDir) {
  try {
    $stage = Join-Path $WorkDir ('specimen_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    $kitDir = Join-Path $stage 'kit'; $pagesDir = Join-Path $stage 'pages'
    foreach ($d in $kitDir, $pagesDir) { New-Item -ItemType Directory -Force $d | Out-Null }
    foreach ($f in 'kit.css', 'kit.js', 'kit-deck.js') { [IO.File]::Copy((Join-Path $script:HdSkill "assets\kit\$f"), (Join-Path $kitDir $f), $true) }
    [IO.File]::Copy($ThemeFile, (Join-Path $kitDir 'theme-preview.css'), $true)
    $html = [IO.File]::ReadAllText((Join-Path $script:HdSkill 'assets\specimen.html'), [Text.Encoding]::UTF8).Replace('../kit/theme-cobalt.css', '../kit/theme-preview.css')
    $page = Join-Path $pagesDir 'specimen.html'
    [IO.File]::WriteAllText($page, $html, (New-Object Text.UTF8Encoding($false)))
    $pages = @([pscustomobject]@{ html = $page; n = 1; canvasW = 1280; canvasH = 720; label = 'theme preview' })
    $run = Invoke-PagePrep -Pages $pages -WorkDir (Join-Path $stage 'out') -SettleMs 1200 -Fonts auto
    if ($run.Fatal) { return '' }
    if ($pages[0].png -and (Test-Path -LiteralPath $pages[0].png)) { [IO.File]::Copy([string]$pages[0].png, $OutPng, $true); return $OutPng }
  } catch { }
  return ''
}

# design.md from the template in assets: every {{KEY}} is replaced by the value of that key (missing keys become '-').
function Write-DesignMd([hashtable]$Values, [string]$Path) {
  $text = [IO.File]::ReadAllText((Join-Path $script:HdSkill 'assets\design-template.md'), [Text.Encoding]::UTF8)
  $text = [regex]::Replace($text, '\{\{([A-Z0-9_]+)\}\}', { param($m) $k = $m.Groups[1].Value; if ($Values.ContainsKey($k) -and "$($Values[$k])" -ne '') { [string]$Values[$k] } else { '-' } })
  [IO.File]::WriteAllText($Path, $text, (New-Object Text.UTF8Encoding($false)))
}

# "v (w), v (w), ..." of a measured list ([{ v, w }]).
function Format-StyleList($List, [int]$Max = 6) {
  $out = @(); foreach ($x in @($List)) { if ($null -eq $x) { continue }; if ($out.Count -ge $Max) { break }; $out += [string]$x.v }
  if ($out.Count -eq 0) { return '-' }
  return ($out -join ', ')
}
