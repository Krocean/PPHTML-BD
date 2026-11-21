param(
  [Parameter(Mandatory)][string]$Html,   # one HTML file, or several separated by |
  [string]$Canvas = '1280x720',          # viewport the page gets, in CSS px: (area width x height in points) x 4/3. Full 16:9 slide = 1280x720.
  [int]$SettleMs = 1500,                 # wait before the picture is taken
  [string]$Engine = 'modern',            # modern (default): the page as the slide show's modern engine (Edge) draws it, exactly as written
                                         # classic: the page converted for the fallback engine (Internet Explorer 11) and checked against its rules
  [string]$Convert = 'auto',             # classic engine: auto = convert modern CSS / JavaScript when the page needs it; always; never
  [string]$Fonts = 'auto',               # auto: embed the fonts the page uses; off
  [string]$Images = 'auto',              # auto: scale large pictures down / store photographs as JPEG; keep
  [string]$Probe = ''                    # a JavaScript expression evaluated in the page after the wait; its value is printed (for debugging)
)
# Quick look at HTML pages the way the slide show will show them, without PowerPoint (a few seconds):
#   - pulls the page's files in and embeds its fonts, as the build does,
#   - loads the page in the engine and reports script errors and missing files,
#   - saves a picture of what the engine drew.
# -Engine classic does the same for the fallback engine and lists what Internet Explorer 11 cannot do.
# Use it while writing a page; html2ppt.ps1 runs the same steps before it builds.
# Exit codes: 0 the pages are fine; 2 they need fixing; 3 a file could not be read; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
. (Join-Path $PSScriptRoot 'lib\pages.ps1')
. (Join-Path $PSScriptRoot 'lib\modern.ps1')
Initialize-Console
$exit = 0
$work = New-WorkDir 'check'
try {
  if ($Canvas -notmatch '^\s*(\d+)\s*[xX*]\s*(\d+)\s*$') { $exit = 3; throw "Canvas must look like 1280x720 (got '$Canvas')." }
  $cw = [int]$Matches[1]; $ch = [int]$Matches[2]
  $Engine = $Engine.ToLowerInvariant(); if ($Engine -eq 'ie11' -or $Engine -eq 'ie') { $Engine = 'classic' }
  if (@('modern', 'classic') -notcontains $Engine) { $exit = 3; throw "Unknown -Engine value '$Engine' (use modern or classic)." }
  $base = (Get-Location).Path
  $pages = @(); $n = 0
  foreach ($h in ($Html -split '\|')) {
    if (-not $h.Trim()) { continue }
    $path = Resolve-UserPath $h $base
    if (-not (Test-Path -LiteralPath $path)) { $exit = 3; throw "HTML file not found: $path" }
    $n++
    $pages += [pscustomobject]@{ html = $path; n = $n; canvasW = $cw; canvasH = $ch; label = 'check only' }
  }
  if ($pages.Count -eq 0) { $exit = 3; throw 'No HTML file given.' }
  $run = $null
  if ($Engine -eq 'modern') {
    $run = Invoke-ModernPrep -Pages $pages -WorkDir $work -SettleMs $SettleMs -Fonts $Fonts -Probe $Probe -Images $Images
    if ($run.Fatal) { $exit = $run.FatalCode; throw $run.Fatal }
    if (-not $run.Ok) { Say "NOTE: Edge cannot be used as the build tool on this PC ($($run.Why)): checking for the classic engine instead."; $Engine = 'classic'; $run = $null }
  }
  if ($Engine -eq 'classic') {
    $run = Invoke-PagePrep -Pages $pages -WorkDir $work -SettleMs $SettleMs -Convert $Convert -Fonts $Fonts -Probe $Probe -Images $Images
    if ($run.Fatal) { $exit = $run.FatalCode; throw $run.Fatal }
    foreach ($x in $run.Warnings) { Say "WARNING: $x" }
  }
  Say $run.Report
  Say ''
  if ($run.Problem -and $Engine -eq 'classic') { $exit = 2; Say 'RESULT: for the classic engine the pages need fixing (ERROR lines above; references\ie11-rules.md has the rewrites). The modern engine runs pages as written.' }
  elseif ($run.Problem) { $exit = 2; Say 'RESULT: the pages need fixing (ERROR lines above): a missing file, or a script error of the page itself.' }
  else { Say 'RESULT: no blocking problems. Look at the picture(s) to confirm the layout; warn lines are worth a look.' }
} catch {
  if ($exit -eq 0) { $exit = 1 }
  Say "FAILED: $($_.Exception.Message)"
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
}
exit $exit
