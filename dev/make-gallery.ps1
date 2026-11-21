param(
  [string]$KitDir = '',        # the kit to draw with (default: the skill's assets\kit). Give a copy when this PC does not let PowerShell read some of its files.
  [string]$Sample = 'kpi',     # the template every theme is shown with
  [switch]$NoTemplates,        # do not draw assets\templates\gallery.png
  [switch]$NoThemes            # do not draw assets\kit\themes.png
)
# Draws the two overview pictures of the design kit (run it after changing the kit, a template or a theme):
#   assets\templates\gallery.png   every template, in the default theme
#   assets\kit\themes.png          one page in every theme
# Needs nothing but the skill itself (PowerShell; PowerPoint is not used). Run one skill command at a time.
$ErrorActionPreference = 'Stop'
$skill = Split-Path -Parent $PSScriptRoot
. (Join-Path $skill 'scripts\lib\common.ps1')
. (Join-Path $skill 'scripts\lib\pages.ps1')
Initialize-Console
$exit = 0
try {
  $work = New-WorkDir 'gallery'
  if (-not $KitDir) { $KitDir = Join-Path $skill 'assets\kit' }
  $index = Read-JsonFile (Join-Path $skill 'assets\templates\index.json')
  Add-Type -AssemblyName System.Drawing
  if (-not ('ImgKit.Compare' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\ImgKit.cs'))) -ReferencedAssemblies System.Drawing }
  $utf8 = New-Object Text.UTF8Encoding($false)
  $stage = Join-Path $work 'src'; $stageKit = Join-Path $stage 'kit'; $stagePages = Join-Path $stage 'pages'
  foreach ($d in $stageKit, $stagePages) { New-Item -ItemType Directory -Force $d | Out-Null }
  $skipped = @()
  foreach ($f in Get-ChildItem -LiteralPath $KitDir -File) {
    if ($f.Extension -ne '.css' -and $f.Extension -ne '.js') { continue }
    try { [IO.File]::Copy($f.FullName, (Join-Path $stageKit $f.Name), $true) } catch { $skipped += $f.Name }
  }
  if ($skipped.Count -gt 0) { Say ("NOTE: this PC does not let PowerShell read: " + ($skipped -join ', ') + "  (left out; give -KitDir a readable copy of the kit to include them)") }

  function Draw($pages, $labels, [string]$out, [int]$cols, [int]$cellW, [string]$name) {
    $run = Invoke-PagePrep -Pages $pages -WorkDir (Join-Path $work $name) -SettleMs 1500 -Fonts auto
    if ($run.Fatal) { throw $run.Fatal }
    $pics = @(); $names = @()
    for ($i = 0; $i -lt @($pages).Count; $i++) {
      if ($pages[$i].png -and (Test-Path -LiteralPath $pages[$i].png)) { $pics += [string]$pages[$i].png; $names += [string]$labels[$i] }
    }
    # pictures made a moment ago are sometimes not readable yet on a PC with security software: try a few times
    for ($try = 1; $try -le 3; $try++) {
      try { [void][ImgKit.Compare]::Sheet([string[]]$pics, $out, $cols, $cellW, [string[]]$names); break }
      catch { if ($try -eq 3) { throw }; Start-Sleep -Seconds 4 }
    }
    Say ("  " + $out + "   (" + $pics.Count + " pictures)")
    if ($run.Problem) { Say '  NOTE: some of these pages reported problems:'; Say $run.Report }
  }

  if (-not $NoTemplates) {
    Say 'Templates ...'
    $pages = @(); $labels = @(); $n = 0
    foreach ($t in @($index.templates)) {
      $src = Join-Path $skill ("assets\templates\$($t.id).html")
      $dst = Join-Path $stagePages ("t-$($t.id).html")
      [IO.File]::WriteAllText($dst, [IO.File]::ReadAllText($src, [Text.Encoding]::UTF8), $utf8)
      $n++; $pages += [pscustomobject]@{ html = $dst; n = $n; canvasW = 1280; canvasH = 720; label = [string]$t.id }; $labels += [string]$t.id
    }
    Draw $pages $labels (Join-Path $skill 'assets\templates\gallery.png') 4 640 'templates'
  }
  if (-not $NoThemes) {
    Say 'Themes ...'
    $sampleHtml = [IO.File]::ReadAllText((Join-Path $skill ("assets\templates\$Sample.html")), [Text.Encoding]::UTF8)
    $pages = @(); $labels = @(); $n = 0; $missing = @()
    foreach ($t in @($index.themes)) {
      $id = [string]$t.id
      if (-not (Test-Path -LiteralPath (Join-Path $stageKit "theme-$id.css"))) { $missing += $id; continue }
      $html = $sampleHtml.Replace('../kit/theme-cobalt.css', "../kit/theme-$id.css")
      $html = [regex]::Replace($html, 'brand: "[^"]*"', ('brand: "' + $id + '"'))
      $dst = Join-Path $stagePages ("theme-$id.html")
      [IO.File]::WriteAllText($dst, $html, $utf8)
      $n++; $pages += [pscustomobject]@{ html = $dst; n = $n; canvasW = 1280; canvasH = 720; label = $id }; $labels += $id
    }
    if ($missing.Count -gt 0) { Say ("  NOTE: not drawn (theme file missing or unreadable): " + ($missing -join ', ')) }
    Draw $pages $labels (Join-Path $skill 'assets\kit\themes.png') 4 640 'themes'
  }
} catch {
  $exit = 1
  Say "FAILED: $($_.Exception.Message)"
  Say $_.ScriptStackTrace
}
exit $exit
