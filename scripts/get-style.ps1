param(
  [Parameter(Mandatory)][string]$From,   # the reference: an HTML file (several: "a.html|b.html"), a web address (https://...),
                                         # a presentation (.pptx .pptm .ppt .potx .ppsx) or a picture (.png .jpg .jpeg .bmp)
  [string]$Out = '',                     # folder for the result (default: style-<name> in the current folder)
  [string]$Name = '',                    # name of the style and of its theme file (default: from the reference)
  [string]$Canvas = '1280x720',          # size of the window an HTML reference is looked at in
  [int]$Slides = 6,                      # pictures taken of a presentation
  [switch]$NoPreview                     # do not draw the preview picture
)
# Reads the LOOK of a reference and writes it down so that a deck can be made in that look:
#   style.json            what was measured, and the few base values the theme is built from
#   theme-<name>.css      a complete theme of the design kit (colours, fonts, corners, shadows)
#   design.md             the look described for people and for Claude: a draft to be corrected against the pictures
#   ref_NN.png            pictures of the reference;  preview.png  the kit drawn with the new theme
# The reference is only read (a presentation is opened read-only, macros off; a web page is only loaded).
# Exit codes: 0 done; 3 the reference cannot be read as given; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\pages.ps1')
. (Join-Path $PSScriptRoot 'lib\style.ps1')
Initialize-Console
$exit = 0
$hdStarted = $false
function NumOf($v, [double]$default) { if ($null -eq $v) { return $default }; $s = [string]$v; if ($s -match '(-?\d+(\.\d+)?)') { return [double]$Matches[1] }; return $default }
function WeightOf($v) { $s = [string]$v; if ($s -eq 'bold') { return 700 }; if ($s -eq 'normal' -or -not $s) { return 400 }; return [int](NumOf $s 400) }
function ModeOf($list) { $best = $null; $n = 0; foreach ($g in (@($list) | Where-Object { $null -ne $_ -and "$_" -ne '' } | Group-Object)) { if ($g.Count -gt $n) { $n = $g.Count; $best = $g.Name } }; return $best }
function Sum-List($acc, $list) { foreach ($x in @($list)) { if ($null -eq $x) { continue }; $k = [string]$x.v; if ($acc.ContainsKey($k)) { $acc[$k] += [double]$x.w } else { $acc[$k] = [double]$x.w } } }
function Top-List($acc, [int]$n) { $o = @(); foreach ($e in ($acc.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $n)) { $o += [pscustomobject][ordered]@{ v = [string]$e.Key; w = [Math]::Round([double]$e.Value, 1) } }; return $o }
try {
  $base = (Get-Location).Path
  $work = New-WorkDir 'style'
  $strings = Read-JsonFile (Join-Path $script:HdSkill 'assets\strings.json')
  $W = $strings.style
  if ($Canvas -notmatch '^\s*(\d+)\s*[xX*]\s*(\d+)\s*$') { $exit = 3; throw "Canvas must look like 1280x720 (got '$Canvas')." }
  $cw = [int]$Matches[1]; $ch = [int]$Matches[2]

  # ---------------------------------------------------------------- what kind of reference
  $sources = @($From -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  if ($sources.Count -eq 0) { $exit = 3; throw 'No reference given.' }
  $kind = ''
  $first = $sources[0]
  if ($first -match '^https?://') { $kind = 'html' }
  else {
    $ext = [IO.Path]::GetExtension($first).ToLowerInvariant()
    if (@('.html', '.htm', '.xhtml') -contains $ext) { $kind = 'html' }
    elseif (@('.pptx', '.pptm', '.ppt', '.potx', '.potm', '.ppsx', '.ppsm') -contains $ext) { $kind = 'deck' }
    elseif (@('.png', '.jpg', '.jpeg', '.bmp', '.gif') -contains $ext) { $kind = 'picture' }
    elseif (@('.md', '.markdown', '.txt') -contains $ext) {
      Say 'A written design description is not measured: read it, then make the theme from the values it names:'
      Say ('  powershell -NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $PSScriptRoot 'new-theme.ps1') + '" -Name <name> -To <deck>\kit -Bg #... -Ink #... -Accent #... -Display "<font>" -Body "<font>" -Radius <px> -Shadow soft|lift|border|none')
      Say 'references\design-from-reference.md says how the parts of such a file map to these values.'
      exit 0
    }
    else { $exit = 3; throw "Do not know how to read '$first' as a reference (HTML file, web address, presentation or picture)." }
  }
  if (-not $Name) {
    $raw = $first; if ($raw -match '^https?://([^/]+)') { $raw = $Matches[1] -replace '^www\.', '' } else { $raw = [IO.Path]::GetFileNameWithoutExtension($raw) }
    $Name = (($raw.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-'))
    if (-not $Name) { $Name = 'ref' }
    if ($Name.Length -gt 24) { $Name = $Name.Substring(0, 24).Trim('-') }
  }
  if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,39}$') { $exit = 3; throw "The name '$Name' must be plain letters, digits, - or _." }
  $Name = $Name.ToLowerInvariant()
  $outDir = $(if ($Out) { Resolve-UserPath $Out $base } else { Join-Path $base "style-$Name" })
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force $outDir | Out-Null }

  $S = @{}                 # base values of the theme
  $D = @{}                 # values for design.md
  $measured = $null; $refs = @(); $fontNotes = @(); $rawLines = @()
  $displayNames = @(); $bodyNames = @()

  if ($kind -eq 'html') {
    # ---------------------------------------------------------------- HTML file(s) or web page: measured in the tool browser
    Say "Reading the look of $($sources.Count) HTML reference(s) in the tool browser ..."
    $edge = Start-EdgeTool $work -NoTools
    if (-not $edge) { $exit = 3; throw "An HTML reference is read with Edge as the tool browser, which cannot be used here: $(Get-EdgeWhyNot)" }
    $js = [IO.File]::ReadAllText((Join-Path $script:HdSkill 'assets\tools\style.js'), [Text.Encoding]::UTF8)
    $acc = @{ bg = @{}; text = @{}; vivid = @{}; fonts = @{}; sizes = @{}; radius = @{}; shadows = @{}; borders = @{} }
    $stacks = @{}; $buttons = @(); $cards = @(); $pageBgs = @(); $h1 = $null; $bodyT = $null; $largest = @(); $vars = $null; $upper = 0.0; $n = 0
    foreach ($src in $sources) {
      $url = $src
      if ($src -notmatch '^https?://') {
        $p = Resolve-UserPath $src $base
        if (-not (Test-Path -LiteralPath $p)) { $exit = 3; throw "Reference not found: $p" }
        $url = ([Uri]$p).AbsoluteUri
      }
      $n++
      $edge.SetViewport($cw, $ch, 1)
      [void]$edge.Navigate($url, 40000)
      [void]$edge.WaitReady(20000)
      Start-Sleep -Milliseconds 1800
      $shot = Join-Path $outDir ("ref_{0:D2}.png" -f $n)
      try { $edge.Screenshot($shot); $refs += $shot } catch { Say "  NOTE: no picture of $src ($($_.Exception.Message))" }
      $one = $null
      try { $one = ($edge.Eval($js, $false, 40000) | ConvertFrom-Json) } catch { Say "  NOTE: $src could not be measured ($($_.Exception.Message))" }
      if (-not $one) { continue }
      foreach ($k in 'bg', 'text', 'vivid', 'fonts', 'sizes', 'radius', 'shadows', 'borders') { Sum-List $acc[$k] $one.$k }
      foreach ($f in @($one.fonts)) { if ($f -and $f.x -and -not $stacks.ContainsKey([string]$f.v)) { $stacks[[string]$f.v] = [string]$f.x } }
      $buttons += @($one.buttons); $cards += @($one.cards); $pageBgs += [string]$one.pageBg
      if (-not $h1) { if ($one.h1) { $h1 = $one.h1 } elseif (@($one.largest).Count -gt 0) { $h1 = @($one.largest)[0] } }
      if (-not $bodyT -and $one.body) { $bodyT = $one.body }
      if ($largest.Count -eq 0) { $largest = @($one.largest) }
      if (-not $vars -and $one.vars) { $vars = $one.vars }
      $upper = [Math]::Max($upper, [double]$one.upperShare)
      Say ("  {0}: {1} characters of text, {2} buttons, {3} cards" -f $src, $one.chars, @($one.buttons).Count, @($one.cards).Count)
    }
    if ($pageBgs.Count -eq 0) { $exit = 3; throw 'Nothing could be measured in the reference.' }
    $S['bg'] = ModeOf $pageBgs
    $textTop = @(Top-List $acc.text 8)
    # the text colour: the colour of the largest heading when it is a strong one, otherwise the strongest of the
    # colours that carry a fair share of the text (the most used one is often the lighter colour of descriptions)
    $inkPick = $null
    if ($h1 -and (Test-Color $h1.c) -and (Get-Contrast $h1.c $S['bg']) -ge 7) { $inkPick = [string]$h1.c }
    if (-not $inkPick -and $textTop.Count -gt 0) {
      $bestC = 0.0
      foreach ($t in $textTop) { if (-not (Test-Color $t.v) -or $t.w -lt $textTop[0].w * 0.2) { continue }; $cr = Get-Contrast $t.v $S['bg']; if ($cr -ge 4.5 -and $cr -gt $bestC -and (Get-Saturation $t.v) -lt 0.5) { $bestC = $cr; $inkPick = [string]$t.v } }
    }
    if ($inkPick) { $S['ink'] = $inkPick }
    $cardBg = ModeOf (@($cards | Where-Object { $_.bg -and $_.bg -ne $S['bg'] } | ForEach-Object { $_.bg }))
    if ($cardBg -and @($cards | Where-Object { $_.bg -eq $cardBg }).Count -ge 2) { $S['surface'] = $cardBg }
    $vivid = @(Top-List $acc.vivid 12 | Where-Object { (Test-Color $_.v) -and (Get-Contrast $_.v $S['bg']) -ge 1.6 })
    $btnColours = @($buttons | Where-Object { $_.bg -and (Get-Saturation $_.bg) -ge 0.25 } | ForEach-Object { $_.bg })
    $btnPlain = @($buttons | Where-Object { $_.bg -and (Get-Saturation $_.bg) -lt 0.25 -and (Get-Contrast $_.bg $S['bg']) -ge 4.5 })
    if ($btnColours.Count -gt 0) { $S['accent'] = ModeOf $btnColours }
    elseif ($vivid.Count -gt 0) { $S['accent'] = $vivid[0].v }
    if ($btnPlain.Count -gt $btnColours.Count -and $btnPlain.Count -ge 1) { $S['button'] = 'ink' }
    $S['colors'] = @($vivid | ForEach-Object { $_.v })
    $fontsTop = @(Top-List $acc.fonts 8)
    if ($h1) { $displayNames = @($h1.f); if ($stacks.ContainsKey([string]$h1.f)) { $displayNames = @($stacks[[string]$h1.f] -split ',') }; $S['displayWeight'] = WeightOf $h1.w
      $ls = [string]$h1.ls; if ($ls -match 'px' -and (NumOf $h1.px 0) -gt 0) { $S['tracking'] = [Math]::Round((NumOf $ls 0) / (NumOf $h1.px 1), 3) } }
    if ($bodyT) { $bodyNames = @($bodyT.f); if ($stacks.ContainsKey([string]$bodyT.f)) { $bodyNames = @($stacks[[string]$bodyT.f] -split ',') } }
    elseif ($fontsTop.Count -gt 0) { $bodyNames = @($fontsTop[0].v) }
    if ($displayNames.Count -eq 0) { $displayNames = $bodyNames }
    $cardR = ModeOf (@($cards | ForEach-Object { [string]$_.r }))
    if ($null -ne $cardR) { $S['radius'] = [double]$cardR } else { $r0 = @(Top-List $acc.radius 4 | Where-Object { $_.v -ne 'pill' }); if ($r0.Count -gt 0) { $S['radius'] = NumOf $r0[0].v 12 } }
    if ($buttons.Count -gt 0) {
      if (@($buttons | Where-Object { $_.pill }).Count * 2 -gt $buttons.Count) { $S['buttonRadius'] = 'pill' } else { $br = ModeOf (@($buttons | ForEach-Object { [string]$_.r })); if ($null -ne $br) { $S['buttonRadius'] = $br } }
    }
    if ($cards.Count -gt 0) {
      $withShadow = @($cards | Where-Object { $_.sh }); $withBorder = @($cards | Where-Object { $_.bd })
      if ($withShadow.Count * 2.5 -gt $cards.Count) { $blur = 0.0; if ([string]$withShadow[0].sh -match '(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px') { $blur = [double]$Matches[3] }; $S['shadow'] = $(if ($blur -ge 12) { 'lift' } else { 'soft' }) }
      elseif ($withBorder.Count * 2.5 -gt $cards.Count) { $S['shadow'] = 'border' }
      else { $S['shadow'] = 'none' }
    }
    if ($upper -gt 0.03) { $S['labelCase'] = 'uppercase' }
    $D['DISPLAY_SIZE'] = $(if ($h1) { "$([Math]::Round((NumOf $h1.px 0)))px" } else { '-' })
    $D['BODY_SIZE'] = $(if ($bodyT) { "$([Math]::Round((NumOf $bodyT.px 0)))px" } else { '-' })
    $D['M_BG'] = Format-StyleList (Top-List $acc.bg 6); $D['M_TEXT'] = Format-StyleList $textTop; $D['M_VIVID'] = Format-StyleList $vivid
    $D['M_FONTS'] = Format-StyleList $fontsTop 5; $D['M_SIZES'] = Format-StyleList (Top-List $acc.sizes 8) 8
    $D['M_RADIUS'] = Format-StyleList (Top-List $acc.radius 5) 5; $D['M_SHADOW'] = Format-StyleList (Top-List $acc.shadows 2) 2; $D['M_BORDER'] = Format-StyleList (Top-List $acc.borders 3) 3
    $D['M_PAD'] = $(if ($cards.Count -gt 0) { [string](ModeOf (@($cards | ForEach-Object { [string]$_.p }))) } else { '-' })
    $D['M_VIEW'] = "$cw x $ch px"
    $measured = [ordered]@{ pageBg = $S['bg']; bg = @(Top-List $acc.bg 10); text = $textTop; vivid = $vivid; fonts = $fontsTop; sizes = @(Top-List $acc.sizes 14); radius = @(Top-List $acc.radius 8)
      shadows = @(Top-List $acc.shadows 5); borders = @(Top-List $acc.borders 6); heading = $h1; body = $bodyT; largest = @($largest); buttons = @($buttons | Select-Object -First 14); cards = @($cards | Select-Object -First 20); cssVariables = $vars }
    $rawLines += 'backgrounds      ' + (Format-StyleList $measured.bg 8)
    $rawLines += 'text colours     ' + (Format-StyleList $textTop 8)
    $rawLines += 'colours with hue ' + (Format-StyleList $vivid 8)
    $rawLines += 'fonts            ' + (Format-StyleList $fontsTop 6)
    $rawLines += 'sizes            ' + (Format-StyleList $measured.sizes 10)
    $rawLines += 'corners          ' + (Format-StyleList $measured.radius 6)
    foreach ($b in @($buttons | Select-Object -First 4)) { $rawLines += ("button           bg {0} text {1} radius {2}px height {3}px weight {4}{5}" -f $b.bg, $b.c, $b.r, $b.h, $b.w, $(if ($b.pill) { ' pill' } else { '' })) }
    foreach ($c in @($cards | Select-Object -First 4)) { $rawLines += ("card             bg {0} radius {1}px border [{2}] padding {3}px size {4}x{5}{6}" -f $c.bg, $c.r, $c.bd, $c.p, $c.w, $c.h, $(if ($c.sh) { ' shadow' } else { '' })) }
    if ($h1) { $rawLines += ("largest heading  {0}px weight {1} font {2} spacing {3} line {4} colour {5}" -f $h1.px, $h1.w, $h1.f, $h1.ls, $h1.lh, $h1.c) }
    if ($bodyT) { $rawLines += ("running text     {0}px weight {1} font {2} line {3} colour {4}" -f $bodyT.px, $bodyT.w, $bodyT.f, $bodyT.lh, $bodyT.c) }
    if ($vars) { $vn = 0; foreach ($p in $vars.PSObject.Properties) { if ($vn -ge 24) { break }; $rawLines += ("css variable     {0}: {1}" -f $p.Name, $p.Value); $vn++ } }
  }
  elseif ($kind -eq 'deck') {
    # ---------------------------------------------------------------- presentation: read by the hidden PowerPoint
    $deck = Resolve-UserPath $first $base
    Say "Reading the look of the presentation (hidden PowerPoint, read-only) ..."
    Start-Hd | Out-Null; $hdStarted = $true
    $run = Invoke-HdChild -Script 'deck_style.ps1' -WorkDir $work -Params @{ deck = $deck; outDir = $outDir; maxSlides = $Slides; scanSlides = 40 } -TimeoutSec 240
    if (-not $run.Ok) { $exit = 3; throw (Format-ChildFailure $run 'Reading the presentation') }
    $m = $run.Result
    $refs = @($m.refs)
    # style.json keeps the names of the pictures only (no folder names of this PC in a file that travels with the deck)
    $m.refs = @($refs | ForEach-Object { Split-Path -Leaf ([string]$_) })
    $k = 720.0 / [double]$m.slideHeight          # points of the slide -> px of the page canvas
    $th = $m.theme
    $bgTop = @($m.bg)
    $bgHex = $null; foreach ($b in $bgTop) { if ([string]$b.v -match '^#') { $bgHex = [string]$b.v; break } elseif ([string]$b.v -match '^gradient (#[0-9A-F]{6})') { $bgHex = $Matches[1]; break } }
    if (-not $bgHex -and $th.light1) { $bgHex = [string]$th.light1 }
    if (-not $bgHex) { $bgHex = '#FFFFFF' }
    $S['bg'] = $bgHex
    # the strongest of the colours that carry a fair share of the text (the most used one is often the softer colour of descriptions)
    $tt = @($m.text); $bestC = 0.0
    foreach ($t in $tt) { if (-not (Test-Color $t.v) -or [double]$t.w -lt [double]$tt[0].w * 0.15) { continue }; $cr = Get-Contrast $t.v $bgHex; if ($cr -ge 4.5 -and $cr -gt $bestC -and (Get-Saturation $t.v) -lt 0.5) { $bestC = $cr; $S['ink'] = [string]$t.v } }
    # colours with a hue: what the slides use (fills by area, text by length), then the accents of the theme
    # (a colour only counts when it stands out from the background: a card a shade lighter than the page is a surface, not an accent)
    $viv = @{}
    foreach ($f in @($m.fill)) { if ((Test-Color $f.v) -and (Get-Saturation $f.v) -ge 0.25 -and (Get-Contrast $f.v $bgHex) -ge 2.0) { $viv[[string]$f.v] = [double]$f.w / 2000.0 * (Get-Saturation $f.v) } }
    foreach ($t in @($m.text)) { if ((Test-Color $t.v) -and (Get-Saturation $t.v) -ge 0.25 -and (Get-Contrast $t.v $bgHex) -ge 2.0) { $kk = [string]$t.v; $add = [double]$t.w * (Get-Saturation $t.v); if ($viv.ContainsKey($kk)) { $viv[$kk] += $add } else { $viv[$kk] = $add } } }
    $vivid = @(Top-List $viv 8)
    $themeAccents = @(); foreach ($a in 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6') { if ($th.$a) { $themeAccents += [string]$th.$a } }
    if ($vivid.Count -gt 0) { $S['accent'] = $vivid[0].v } elseif ($themeAccents.Count -gt 0) { $S['accent'] = $themeAccents[0] }
    $S['colors'] = @(@($vivid | ForEach-Object { $_.v }) + $themeAccents)
    $surf = $null; foreach ($f in @($m.fill)) { if ((Test-Color $f.v) -and $f.v -ne $bgHex) { $cr = Get-Contrast $f.v $bgHex; if ($cr -gt 1.03 -and $cr -lt 1.8) { $surf = [string]$f.v; break } } }
    if ($surf) { $S['surface'] = $surf }
    $tf = @($m.titleFont); $bf = @($m.bodyFont); $ts = @($m.titleSize); $bs = @($m.bodySize)
    $sz = @($m.sizes)
    if ($tf.Count -eq 0 -and $sz.Count -gt 0) {
      # no title placeholders (text boxes only): the largest size that carries a fair share of the text is the headline,
      # the size with the most text among the rest is the running text
      $total = 0.0; foreach ($x in $sz) { $total += [double]$x.chars }
      $head = $null; foreach ($x in $sz) { if ([double]$x.chars -ge $total * 0.04) { $head = $x; break } }
      if (-not $head) { $head = $sz[0] }
      $rest = @($sz | Where-Object { $_.pt -lt $head.pt } | Sort-Object -Property chars -Descending)
      $tf = @([pscustomobject]@{ v = [string]$head.font; w = $head.chars }); $ts = @([pscustomobject]@{ v = [string]$head.pt; w = $head.chars })
      $S['displayWeight'] = $(if ([double]$head.boldShare -ge 0.5) { 700 } else { 400 })
      if ($rest.Count -gt 0) { $bf = @([pscustomobject]@{ v = [string]$rest[0].font; w = $rest[0].chars }); $bs = @([pscustomobject]@{ v = [string]$rest[0].pt; w = $rest[0].chars }) }
    } elseif ([int]$m.titleChars -gt 0) { $S['displayWeight'] = $(if ([double]$m.titleBoldShare -ge 0.5) { 700 } else { 400 }) }
    if ($tf.Count -gt 0) { $displayNames = @($tf | ForEach-Object { [string]$_.v }) } elseif ($m.themeFonts.majorEast -or $m.themeFonts.majorLatin) { $displayNames = @([string]$m.themeFonts.majorLatin, [string]$m.themeFonts.majorEast) }
    if ($bf.Count -gt 0) { $bodyNames = @($bf | ForEach-Object { [string]$_.v }) } elseif ($m.themeFonts.minorEast -or $m.themeFonts.minorLatin) { $bodyNames = @([string]$m.themeFonts.minorLatin, [string]$m.themeFonts.minorEast) }
    if ($displayNames.Count -eq 0) { $displayNames = $bodyNames }
    $rt = @($m.radius)
    # corner radius was measured in px of the slide at 100 % (points x 4/3): bring it to the 720 px page canvas
    if ($rt.Count -gt 0) { $rv = [string]$rt[0].v; if ($rv -eq 'round') { $S['radius'] = 24 } else { $S['radius'] = [Math]::Round((NumOf $rv 0) * 720.0 / ([double]$m.slideHeight * 4.0 / 3.0)) } }
    $S['shadow'] = $(if ([double]$m.shadowShare -ge 0.3) { 'soft' } elseif (@($m.line).Count -gt 0) { 'border' } else { 'none' })
    $D['DISPLAY_SIZE'] = $(if ($ts.Count -gt 0) { "$([Math]::Round((NumOf $ts[0].v 0) * $k))px ($($ts[0].v)pt)" } else { '-' })
    $D['BODY_SIZE'] = $(if ($bs.Count -gt 0) { "$([Math]::Round((NumOf $bs[0].v 0) * $k))px ($($bs[0].v)pt)" } else { '-' })
    $D['M_BG'] = Format-StyleList $bgTop; $D['M_TEXT'] = Format-StyleList $m.text; $D['M_VIVID'] = Format-StyleList $vivid
    $D['M_FONTS'] = (Format-StyleList $tf 3) + ' / ' + (Format-StyleList $bf 3)
    $D['M_SIZES'] = (Format-StyleList $ts 3) + ' pt / ' + (Format-StyleList $bs 5) + ' pt'
    $D['M_RADIUS'] = Format-StyleList $rt 4; $D['M_SHADOW'] = "$([Math]::Round([double]$m.shadowShare * 100))%"; $D['M_BORDER'] = Format-StyleList $m.line 3
    $D['M_PAD'] = '-'; $D['M_VIEW'] = "$([Math]::Round([double]$m.slideWidth)) x $([Math]::Round([double]$m.slideHeight)) pt ($($m.slides) slides)"
    $measured = $m
    $rawLines += 'slide            ' + $D['M_VIEW']
    $rawLines += 'theme colours    ' + ((@($th.PSObject.Properties | ForEach-Object { "$($_.Name) $($_.Value)" })) -join ', ')
    $rawLines += 'theme fonts      ' + ((@($m.themeFonts.PSObject.Properties | ForEach-Object { "$($_.Name) $($_.Value)" })) -join ', ')
    $rawLines += 'backgrounds      ' + (Format-StyleList $bgTop 6)
    $rawLines += 'text colours     ' + (Format-StyleList $m.text 8)
    $rawLines += 'fills (by area)  ' + (Format-StyleList $m.fill 10)
    $rawLines += 'lines            ' + (Format-StyleList $m.line 5)
    $rawLines += 'title font/size  ' + (Format-StyleList $tf 3) + ' / ' + (Format-StyleList $ts 4) + ' pt'
    $rawLines += 'text font/size   ' + (Format-StyleList $bf 3) + ' / ' + (Format-StyleList $bs 6) + ' pt'
    $rawLines += 'corners (px)     ' + (Format-StyleList $rt 5)
    $rawLines += "shadows on boxes $([Math]::Round([double]$m.shadowShare * 100))%   pictures $($m.pictures)"
  }
  else {
    # ---------------------------------------------------------------- picture: its main colours only
    $pic = Resolve-UserPath $first $base
    if (-not (Test-Path -LiteralPath $pic)) { $exit = 3; throw "Reference not found: $pic" }
    Add-Type -AssemblyName System.Drawing
    if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
    $pal = ''
    try { $pal = [ImgKit.Compare]::Palette($pic, 10) }
    catch { $exit = 3; throw "ACCESS-DENIED or unreadable picture: PowerShell could not read '$pic' on this PC. Copy it to a folder under '$(Join-Path (Get-HdRoot) 'in')' with your own file tools and give that copy." }
    $cols = @(); foreach ($e in ($pal -split ';')) { if ($e -match '^([0-9A-F]{6}):([\d.]+)$') { $cols += [pscustomobject]@{ v = '#' + $Matches[1]; w = [double]$Matches[2] } } }
    if ($cols.Count -eq 0) { $exit = 3; throw 'No colours could be read from the picture.' }
    try { $ref = Join-Path $outDir 'ref_01.png'; $img = [Drawing.Image]::FromFile($pic); try { $img.Save($ref, [Drawing.Imaging.ImageFormat]::Png) } finally { $img.Dispose() }; $refs += $ref } catch {}
    $S['bg'] = $cols[0].v
    $bestInk = $null; $bestC = 0.0
    foreach ($c in $cols) { $cr = Get-Contrast $c.v $S['bg']; if ($cr -gt $bestC -and (Get-Saturation $c.v) -lt 0.45) { $bestC = $cr; $bestInk = $c.v } }
    # (thin text blends with its background when a picture is scaled down: only a clearly strong colour is taken as the text colour)
    if ($bestInk -and $bestC -ge 7) { $S['ink'] = $bestInk }
    $pageBg = [string]$S['bg']
    $vivid = @($cols | Where-Object { (Get-Saturation $_.v) -ge 0.3 -and $_.w -ge 0.2 -and (Get-Contrast $_.v $pageBg) -ge 2.0 } | Sort-Object -Property @{ Expression = { $_.w * (Get-Saturation $_.v) } } -Descending)
    if ($vivid.Count -gt 0) { $S['accent'] = $vivid[0].v }
    $S['colors'] = @($vivid | ForEach-Object { $_.v })
    $D['DISPLAY_SIZE'] = '-'; $D['BODY_SIZE'] = '-'
    $D['M_BG'] = Format-StyleList $cols 4; $D['M_TEXT'] = '-'; $D['M_VIVID'] = Format-StyleList $vivid; $D['M_FONTS'] = '-'; $D['M_SIZES'] = '-'; $D['M_RADIUS'] = '-'; $D['M_SHADOW'] = '-'; $D['M_BORDER'] = '-'; $D['M_PAD'] = '-'; $D['M_VIEW'] = '-'
    $measured = [ordered]@{ palette = $cols }
    $rawLines += 'colours (share of the picture %)  ' + ((@($cols | ForEach-Object { "$($_.v) $($_.w)" })) -join ', ')
    Say '  A picture gives colours only: fonts, corners and layout have to be read from it by eye (design.md, sections 3 to 5).'
  }

  # ---------------------------------------------------------------- fonts that can be embedded here
  $ds = Get-StyleFontStack $displayNames 'sans'
  $bsk = 'sans'; if ($ds.serif -and $bodyNames.Count -eq 0) { $bsk = 'serif' }
  $bst = Get-StyleFontStack $(if ($bodyNames.Count -gt 0) { $bodyNames } else { $displayNames }) $bsk
  $S['display'] = $ds.stack; $S['body'] = $bst.stack
  $fontNotes = @(@($ds.notes) + @($bst.notes) | Select-Object -Unique)

  # ---------------------------------------------------------------- theme, description, preview
  $made = New-ThemeTokens $S
  $t = $made.tokens
  $themeFile = Join-Path $outDir "theme-$Name.css"
  # local files are named without their folder: this text goes into the theme file, design.md and style.json,
  # which travel with the deck (and the theme into every page), so a path with the user's folder names must not be in it
  $srcText = (@($sources | ForEach-Object { if ($_ -match '^https?://') { $_ } else { [IO.Path]::GetFileName($_) } }) -join ', ')
  Write-ThemeCss $t $Name ("the look of " + ($srcText -replace '[^\x20-\x7E]', '?')) $themeFile
  $baseOut = [ordered]@{}; foreach ($kk in ($S.Keys | Sort-Object)) { $baseOut[$kk] = $S[$kk] }
  $styleJson = [ordered]@{ name = $Name; source = $srcText; kind = $kind; made = (Get-Date).ToString('yyyy-MM-dd'); base = $baseOut; tokens = $t; notes = @(@($made.notes) + $fontNotes); measured = $measured }
  [IO.File]::WriteAllText((Join-Path $outDir 'style.json'), ($styleJson | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
  $D['NAME'] = $Name; $D['SOURCE'] = $srcText; $D['DATE'] = (Get-Date).ToString('yyyy-MM-dd')
  $D['REFS'] = $(if ($refs.Count -gt 0) { (@($refs | ForEach-Object { '`' + (Split-Path -Leaf $_) + '`' }) -join ', ') } else { '-' })
  $D['BG'] = $t['bg']; $D['SURFACE'] = $t['surface']; $D['INK'] = $t['ink']; $D['INK2'] = $t['ink2']; $D['LINE'] = $t['line']; $D['ACCENT'] = $t['accent']; $D['ACCENT_INK'] = $t['accent-ink']
  $D['CHART'] = (@(1..6 | ForEach-Object { '`' + $t["c$_"] + '`' } | Select-Object -Unique) -join ' ')
  $D['MODE'] = $(if ($made.dark) { [string]$W.dark } else { [string]$W.light })
  $D['DISPLAY_FONT'] = $t['display']; $D['DISPLAY_WEIGHT'] = $t['display-weight']; $D['BODY_FONT'] = $t['body']; $D['STRONG_WEIGHT'] = $t['strong-weight']
  $D['TRACKING'] = $t['display-tracking']; $D['LABEL_CASE'] = $(if ($t['label-case'] -eq 'uppercase') { [string]$W.caseUpper } else { [string]$W.caseNone })
  $D['FONT_NOTES'] = $(if ($fontNotes.Count -gt 0) { ($fontNotes -join '; ') } else { [string]$W.fontKept })
  $D['RADIUS'] = $t['radius']; $D['BUTTON_RADIUS'] = $(if ($t['radius-btn'] -eq '999px') { [string]$W.pill } else { $t['radius-btn'] })
  $shadowWord = [string]$W.shadowSoft
  if ($t['card-border'] -like '1px*') { $shadowWord = [string]$W.shadowBorder } elseif ($t['shadow'] -like '0 0 0 0*') { $shadowWord = [string]$W.shadowNone } elseif ([string]$S['shadow'] -eq 'lift') { $shadowWord = [string]$W.shadowLift }
  $D['SHADOW'] = $shadowWord
  $D['BUTTON'] = $(if ([string]$S['button'] -eq 'ink') { [string]$W.buttonInk } else { [string]$W.buttonAccent }) + " ($($t['press']))"
  $D['THEME_FILE'] = "theme-$Name.css"
  $D['RAW'] = '```' + "`n" + ($rawLines -join "`n") + "`n" + '```'
  Write-DesignMd $D (Join-Path $outDir 'design.md')

  Say ''
  Say "Style '$Name' written to: $outDir"
  Say ("  theme-$Name.css   page {0}  card {1}  text {2}  accent {3}  {4}" -f $t['bg'], $t['surface'], $t['ink'], $t['accent'], $(if ($made.dark) { 'dark' } else { 'light' }))
  Say ("                    headlines {0} / {1}   text {2}   corners {3}" -f $t['display'], $t['display-weight'], $t['body'], $t['radius'])
  Say '  design.md         the look in words: a draft - correct it against the pictures'
  Say '  style.json        everything that was measured'
  foreach ($r in $refs) { Say "  $(Split-Path -Leaf $r)         picture of the reference: $r" }
  foreach ($n2 in (@($made.notes) + $fontNotes)) { if ($n2) { Say "  NOTE: $n2" } }
  if (-not $NoPreview) {
    $png = New-StylePreview $themeFile (Join-Path $outDir 'preview.png') $work
    if ($png) { Say "  preview.png       the kit drawn with this theme: $png" } else { Say '  NOTE: the preview picture could not be drawn.' }
  }
  Say ''
  Say 'Next: look at the reference picture(s) and preview.png side by side, correct design.md and the theme'
  Say ('  (new-theme.ps1 -From "' + (Join-Path $outDir 'style.json') + '" -Name ' + $Name + ' -To "' + $outDir + '" -Accent #... -Radius ... changes single values),')
  Say ('  then start the deck: new-deck.ps1 -To <folder> -Design "' + $outDir + '" -Blank <pages>')
} catch {
  if ($exit -eq 0) { $exit = 1 }
  Say "FAILED: $($_.Exception.Message)"
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
} finally {
  try { Stop-EdgeTool } catch {}
  if ($hdStarted) { try { Stop-Hd } catch {} }
}
exit $exit
