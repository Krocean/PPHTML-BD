param(
  [Parameter(Mandatory)][string]$Deck,   # presentation to convert (.pptx, .pptm, .ppt, .ppsx ...)
  [string]$Out,                          # result (.html). Default: next to the presentation, same name.
  [int]$Width = 1600,                    # width of the slide pictures in pixels (height follows the slide shape)
  [switch]$Jpg,                          # JPEG pictures instead of PNG: much smaller for photo-heavy decks, slightly softer text
  [switch]$KeepHidden                    # leave the hidden PowerPoint running afterwards
)
# PowerPoint -> HTML: one self-contained .html file that shows the presentation in any current browser
# (no PowerPoint, no macros, no internet). Arrow keys / click / swipe move between slides; G = slide list,
# N = speaker notes, F = full screen. Hyperlinks keep working. If the presentation carries live HTML pages
# (built with html2ppt.ps1), those pages run in place.
# Exit codes: 0 done; 3 the presentation / this PC cannot be used as asked; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\webdeck.ps1')
. (Join-Path $PSScriptRoot 'lib\edge.ps1')
. (Join-Path $PSScriptRoot 'lib\webcheck.ps1')
Initialize-Console
$exit = 0
$hdStarted = $false
$work = New-WorkDir 'ppt2html'
$Res = [ordered]@{ ok = $false; out = ''; webCheck = $null; workDir = $work; slides = 0; livePages = 0; bytes = 0; warnings = @() }
try {
  $base = (Get-Location).Path
  $deckPath = Resolve-UserPath $Deck $base
  if (-not (Test-Path -LiteralPath $deckPath)) { $exit = 3; throw "Presentation not found: $deckPath" }
  if (-not $Out) { $Out = [IO.Path]::ChangeExtension($deckPath, '.html') }
  $outPath = Resolve-UserPath $Out $base
  if ([IO.Path]::GetExtension($outPath).ToLowerInvariant() -notin '.html', '.htm') { $outPath = $outPath + '.html' }
  Say "[1/2] Reading the slides (hidden PowerPoint) ..."
  try { Start-Hd | Out-Null; $hdStarted = $true } catch { $exit = 3; throw }
  $fmt = if ($Jpg) { 'jpg' } else { 'png' }
  $wd = New-WebDeck -Source $deckPath -OutFile $outPath -WorkDir $work -Width $Width -ImageFormat $fmt -Title ([IO.Path]::GetFileNameWithoutExtension($deckPath))
  Say "[2/2] Saved."
  foreach ($x in $wd.warnings) { $Res.warnings += $x; Say "  WARNING: $x" }
  $Res.out = $wd.out; $Res.slides = $wd.slides; $Res.livePages = $wd.livePages; $Res.bytes = $wd.bytes
  Say ("  {0}   ({1} slides, {2} live HTML page(s), {3:N1} MB)" -f $wd.out, $wd.slides, $wd.livePages, ($wd.bytes / 1MB))
  $shots = Join-Path $work 'shots'; New-Item -ItemType Directory -Force $shots | Out-Null
  $wc = Test-WebDeck -File $wd.out -WorkDir $work -ShotsDir $shots
  $Res.webCheck = [ordered]@{ checked = $wc.checked; slides = $wc.slides; livePages = $wc.live; liveOk = $wc.liveOk; errors = @($wc.errors); pictures = @($wc.shots) }
  if (-not $wc.checked) { Say "  not checked in a browser: $($wc.why)" }
  else {
    Say "  checked in Edge: $($wc.slides) slide(s) open$(if ($wc.live) { ", $($wc.liveOk) of $($wc.live) live page(s) running" } else { '' })$(if (@($wc.errors).Count -eq 0) { ', no script errors' } else { '' })"
    foreach ($e in @($wc.errors)) { $Res.warnings += "Web version: $e"; Say "  WARNING: Web version: $e" }
    foreach ($s in @($wc.shots | Select-Object -First 3)) { Say "  picture: $s" }
  }
  $Res.ok = $true
  Say ''
  Say 'DONE.'
} catch {
  if ($exit -eq 0) { $exit = 1 }
  $Res.error = $_.Exception.Message
  Say "FAILED: $($_.Exception.Message)"
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
} finally {
  if ($hdStarted -and -not $KeepHidden) { try { [void](Stop-Hd) } catch {} }
  try { Write-JsonFile (Join-Path $work 'result.json') $Res } catch {}
  Say "Details: $(Join-Path $work 'result.json')"
}
exit $exit
