param(
  [string]$Name = 'custom',          # name of the theme: the file becomes theme-<name>.css
  [string]$To = '',                  # folder to write it to (a deck's kit folder, or any folder). Default: the current folder
  [string]$From = '',                # style.json of get-style.ps1 to start from (its values; the options below override them)
  [string]$Bg = '', [string]$Surface = '', [string]$Ink = '', [string]$Accent = '',   # colours, #RRGGBB
  [string]$Colors = '',              # more chart colours, in order: "#0EA5A4,#F59E0B"
  [string]$Display = '',             # headline font(s): "Inter, Pretendard"  (what is not on this PC or in the font pack is replaced)
  [string]$Body = '',                # text font(s)
  [int]$DisplayWeight = 0,           # 100-900
  [int]$StrongWeight = 0,            # weight of bold text, 400-900
  [string]$Tracking = '',            # letter-spacing of headlines in em: -0.02
  [string]$Radius = '',              # corner of cards in px
  [string]$ButtonRadius = '',        # px, or pill
  [string]$Shadow = '',              # soft | lift | border | none
  [string]$Button = '',              # accent | ink : colour of the main button
  [string]$LabelCase = '',           # none | uppercase
  [string]$Note = '',                # one line for the head of the file (where the look comes from)
  [switch]$NoPreview                 # do not draw the preview picture
)
# Makes a complete theme of the design kit from a few values: the rest (tints, text shades, lines, hover, status
# and chart colours, shadows) is worked out, contrast is checked, and a preview picture is drawn with it.
# Use it to give a deck the look of a reference (after get-style.ps1), of a design.md, or of a company style.
# Exit codes: 0 done; 3 cannot be done as asked; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\pages.ps1')
. (Join-Path $PSScriptRoot 'lib\style.ps1')
Initialize-Console
$exit = 0
try {
  if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,39}$') { $exit = 3; throw "The theme name '$Name' must be plain letters, digits, - or _ (it becomes a file name)." }
  $base = (Get-Location).Path
  $S = @{}
  if ($From) {
    $fp = Resolve-UserPath $From $base
    if (-not (Test-Path -LiteralPath $fp)) { $exit = 3; throw "Not found: $fp" }
    $j = Read-JsonFile $fp
    if ($j.base) { foreach ($p in $j.base.PSObject.Properties) { $S[$p.Name] = $p.Value } }
  }
  foreach ($pair in @(@('bg', $Bg), @('surface', $Surface), @('ink', $Ink), @('accent', $Accent))) {
    if ($pair[1]) { if (-not (Test-Color $pair[1])) { $exit = 3; throw "'$($pair[1])' is not a colour like #1A2B3C." }; $S[$pair[0]] = $pair[1] }
  }
  if ($Colors) { $S['colors'] = @($Colors -split '[,; ]+' | Where-Object { $_ }) }
  $fontNotes = @()
  if ($Display) { $fs = Get-StyleFontStack @($Display -split ',') 'sans'; $S['display'] = $fs.stack; $fontNotes += $fs.notes }
  if ($Body) { $fs = Get-StyleFontStack @($Body -split ',') 'sans'; $S['body'] = $fs.stack; $fontNotes += $fs.notes }
  if ($DisplayWeight -gt 0) { $S['displayWeight'] = $DisplayWeight }
  if ($StrongWeight -gt 0) { $S['strongWeight'] = $StrongWeight }
  if ($Tracking) { $S['tracking'] = [double]($Tracking -replace 'em$', '') }
  if ($Radius -ne '') { $S['radius'] = [double]($Radius -replace 'px$', '') }
  if ($ButtonRadius -ne '') { $S['buttonRadius'] = ($ButtonRadius -replace 'px$', '') }
  if ($Shadow) { if (@('soft', 'lift', 'border', 'none') -notcontains $Shadow.ToLowerInvariant()) { $exit = 3; throw "Shadow must be soft, lift, border or none." }; $S['shadow'] = $Shadow.ToLowerInvariant() }
  if ($Button) { $S['button'] = $Button.ToLowerInvariant() }
  if ($LabelCase) { $S['labelCase'] = $LabelCase.ToLowerInvariant() }

  $made = New-ThemeTokens $S
  $dir = $base; if ($To) { $dir = Resolve-UserPath $To $base }
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
  $file = Join-Path $dir ("theme-$($Name.ToLowerInvariant()).css")
  $head = $Note; if (-not $head) { $head = 'a custom look' }
  Write-ThemeCss $made.tokens $Name.ToLowerInvariant() $head $file
  Say "Theme: $file"
  $t = $made.tokens
  Say ("  page {0}  card {1}  text {2}  accent {3} (text on it {4})  {5}" -f $t['bg'], $t['surface'], $t['ink'], $t['accent'], $t['accent-ink'], $(if ($made.dark) { 'dark' } else { 'light' }))
  Say ("  chart colours: " + (@(1..6 | ForEach-Object { $t["c$_"] }) -join ' '))
  Say ("  headlines: {0}  weight {1}  spacing {2}" -f $t['display'], $t['display-weight'], $t['display-tracking'])
  Say ("  text: {0}   corners: {1} / buttons {2}   card edge: {3}" -f $t['body'], $t['radius'], $t['radius-btn'], $(if ($t['card-border'] -like '1px*') { 'hairline' } elseif ($t['shadow'] -like '0 0 0 0*') { 'none' } else { 'shadow' }))
  Say ("  contrast: text {0}:1, accent {1}:1 on the page" -f [Math]::Round((Get-Contrast $t['ink'] $t['bg']), 1), [Math]::Round((Get-Contrast $t['accent'] $t['bg']), 1))
  foreach ($n in (@($made.notes) + @($fontNotes))) { if ($n) { Say "  NOTE: $n" } }
  if (-not $NoPreview) {
    $work = New-WorkDir 'theme'
    $png = New-StylePreview $file (Join-Path $dir ("theme-$($Name.ToLowerInvariant()).preview.png")) $work
    if ($png) { Say "  preview (every part of the kit drawn with this theme): $png" } else { Say '  NOTE: the preview picture could not be drawn (the theme file itself is fine).' }
  }
  Say ''
  Say "Use it: link ../kit/theme-$($Name.ToLowerInvariant()).css at the top of each page (new-deck.ps1 -Theme `"$file`" does that for a new deck)."
} catch {
  if ($exit -eq 0) { $exit = 1 }
  Say "FAILED: $($_.Exception.Message)"
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
} finally { try { Stop-EdgeTool } catch {} }
exit $exit
