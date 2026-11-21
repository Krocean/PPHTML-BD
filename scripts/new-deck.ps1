param(
  [string]$To,                       # folder to create the deck project in (made if missing)
  [int]$Blank = 0,                   # this many empty pages to design freely (the usual way to start)
  [string]$Titles = '',              # titles of the empty pages, separated by |  (they also name the entries of the header menu)
  [string]$Design = '',              # a style folder made by get-style.ps1: its theme, design.md and reference pictures go into the deck
  [string]$Theme = '',               # a theme of the kit (see -List) or the path of a theme file (theme-xxx.css). Default: cobalt
  [string]$Recipe = '',              # ready-made example pages in a fitting order: review, pitch, keynote, report, workshop (see -List)
  [string]$Pages = '',               # example pages by name, in slide order (see -List); "blank" stands for an empty page: "cover,blank,blank,closing"
  [string]$Name = '',                # name of the presentation (file name and the brand in the page header). Default: the folder name
  [switch]$NoNav,                    # pages without the site header menu
  [switch]$List,                     # only list the example pages, recipes and themes
  [switch]$Force                     # overwrite page files that already exist in the folder
)
# Starts a deck project: a folder with the design kit (theme, styles, script), numbered pages wired to each
# other, and a job file. The pages are empty canvases to design freely (-Blank), example pages to rework
# (-Recipe / -Pages), or a mix. With -Design the deck takes the look that get-style.ps1 read from a reference.
# After that: write the pages, then build with  html2ppt.ps1 -Job "<folder>\job.json"
# Exit codes: 0 done; 3 cannot be done as asked; 1 unexpected error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
Initialize-Console
$exit = 0
try {
  $skill = $script:HdSkill
  $index = Read-JsonFile (Join-Path $skill 'assets\templates\index.json')
  $strings = Read-JsonFile (Join-Path $skill 'assets\strings.json')
  if ($List -or -not $To) {
    Say 'Empty pages to design freely:  -Blank <number>  [-Titles "a|b|c"]   (or "blank" inside -Pages)'
    Say 'The look of a reference:       -Design "<style folder of get-style.ps1>"   or   -Theme "<theme file>"'
    Say ''
    Say 'Example pages (use with -Pages, in the order of the slides; rework them, they are starting points):'
    foreach ($t in @($index.templates)) { Say ("  {0,-12} {1} - {2}" -f $t.id, $t.title, $t.what) }
    Say ''
    Say 'Recipes (use with -Recipe: example pages in a fitting order, with a fitting theme):'
    foreach ($t in @($index.recipes)) { Say ("  {0,-12} {1} - {2}" -f $t.id, $t.title, $t.what); Say ("               pages: {0}   theme: {1}" -f $t.pages, $t.theme) }
    Say ''
    Say 'Themes of the kit (use with -Theme):'
    foreach ($t in @($index.themes)) { Say ("  {0,-12} {1}" -f $t.id, $t.what) }
    Say ''
    $gal = Join-Path $skill 'assets\templates\gallery.png'
    if (Test-Path -LiteralPath $gal) { Say ('Pictures of every example page: ' + $gal) }
    $thp = Join-Path $skill 'assets\kit\themes.png'
    if (Test-Path -LiteralPath $thp) { Say ('Pictures of every theme:        ' + $thp) }
    Say 'How to take the look of a reference: references\design-from-reference.md.  How to order and word a deck: references\storytelling.md'
    if (-not $List) { Say ''; Say 'Give -To "<folder>" to create a project.' }
    exit 0
  }
  $here = (Get-Location).Path
  if ($Recipe) {
    $r = @($index.recipes) | Where-Object { $_.id -eq $Recipe.ToLowerInvariant() } | Select-Object -First 1
    if (-not $r) { $exit = 3; throw "Unknown recipe '$Recipe' (see -List)." }
    if (-not $Pages) { $Pages = [string]$r.pages }
    if (-not $Theme -and -not $Design) { $Theme = [string]$r.theme }
  }

  # ---------------------------------------------------------------- the look: a style folder, a theme file, or a theme of the kit
  $customTheme = ''; $designDir = ''
  if ($Design) {
    $designDir = Resolve-UserPath $Design $here
    if (-not (Test-Path -LiteralPath $designDir)) { $exit = 3; throw "Style folder not found: $designDir" }
    $found = @(Get-ChildItem -LiteralPath $designDir -Filter 'theme-*.css' -File | Sort-Object LastWriteTime -Descending)
    if ($found.Count -eq 0 -and -not $Theme) { $exit = 3; throw "No theme-*.css in '$designDir' (get-style.ps1 or new-theme.ps1 writes one)." }
    if ($found.Count -gt 0 -and -not $Theme) { $customTheme = $found[0].FullName }
  }
  if ($Theme -and ($Theme -match '\.css$')) {
    $customTheme = Resolve-UserPath $Theme $here
    if (-not (Test-Path -LiteralPath $customTheme)) { $exit = 3; throw "Theme file not found: $customTheme" }
  }
  $themeName = 'cobalt'
  if ($customTheme) {
    $themeName = ([IO.Path]::GetFileNameWithoutExtension($customTheme) -replace '^theme-', '').ToLowerInvariant()
    if ($themeName -notmatch '^[a-z0-9][a-z0-9_-]*$') { $exit = 3; throw "The theme file must be named theme-<plain name>.css (got '$customTheme')." }
  } elseif ($Theme) {
    $themeName = $Theme.ToLowerInvariant()
    if (-not (Test-Path -LiteralPath (Join-Path $skill "assets\kit\theme-$themeName.css"))) { $exit = 3; throw "Unknown theme '$Theme' (see -List), and not a theme file." }
  }

  # ---------------------------------------------------------------- the pages
  $plan = @()
  if ($Pages) {
    foreach ($id in @($Pages -split '[,| ]+' | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() })) {
      if ($id -eq 'blank') { $plan += [pscustomobject]@{ id = 'blank'; title = '' }; continue }
      $t = @($index.templates) | Where-Object { $_.id -eq $id } | Select-Object -First 1
      if (-not $t) { $exit = 3; throw "Unknown example page '$id' (see -List; 'blank' is an empty page)." }
      $plan += [pscustomobject]@{ id = [string]$t.id; title = [string]$t.title }
    }
  }
  for ($i = 0; $i -lt $Blank; $i++) { $plan += [pscustomobject]@{ id = 'blank'; title = '' } }
  if ($plan.Count -eq 0) { for ($i = 0; $i -lt 3; $i++) { $plan += [pscustomobject]@{ id = 'blank'; title = '' } } }
  if ($plan.Count -gt 30) { $exit = 3; throw 'A deck takes at most 30 pages.' }
  $given = @(); if ($Titles) { $given = @($Titles -split '\|' | ForEach-Object { $_.Trim() }) }
  $b = 0; $n = 0
  foreach ($p in $plan) {
    $n++
    if ($p.id -eq 'blank') {
      if ($b -lt $given.Count -and $given[$b]) { $p.title = $given[$b] } else { $p.title = ([string]$strings.newDeck.blankTitle).Replace('{n}', [string]$n) }
      $b++
    }
  }

  $root = Resolve-UserPath $To $here
  if (-not $Name) { $Name = Split-Path -Leaf $root }
  $kitDir = Join-Path $root 'kit'; $pagesDir = Join-Path $root 'pages'
  foreach ($d in $root, $kitDir, $pagesDir) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force $d | Out-Null } }
  # a file this PC does not let PowerShell read (security software) is left out; the deck needs only its own theme and the kit itself
  $skipped = @()
  foreach ($f in Get-ChildItem -LiteralPath (Join-Path $skill 'assets\kit') -File) {
    if ($f.Extension -eq '.png') { continue }
    try { [IO.File]::Copy($f.FullName, (Join-Path $kitDir $f.Name), $true) } catch { $skipped += $f.Name }
  }
  if ($customTheme) { try { [IO.File]::Copy($customTheme, (Join-Path $kitDir "theme-$themeName.css"), $true) } catch { $skipped += "theme-$themeName.css" } }
  $need = @(@("theme-$themeName.css", 'kit.css', 'kit.js', 'kit-deck.js') | Where-Object { ($skipped -contains $_) -and -not (Test-Path -LiteralPath (Join-Path $kitDir $_)) })
  if ($need.Count -gt 0) {
    $exit = 3
    throw ("This PC does not let PowerShell read " + ($need -join ', ') + " (security software is the usual cause). Pick another theme, or copy that file into '$kitDir' with your own file tool and run this again.")
  }
  $designOut = ''
  if ($designDir) {
    $designOut = Join-Path $root 'design'
    if (-not (Test-Path -LiteralPath $designOut)) { New-Item -ItemType Directory -Force $designOut | Out-Null }
    foreach ($f in Get-ChildItem -LiteralPath $designDir -File) {
      if (@('.md', '.json', '.png', '.css') -notcontains $f.Extension.ToLowerInvariant()) { continue }
      try { [IO.File]::Copy($f.FullName, (Join-Path $designOut $f.Name), $true) } catch { $skipped += ('design\' + $f.Name) }
    }
  }

  # header menu: one entry per page, pointing at its slide number
  Add-Type -AssemblyName System.Web.Extensions
  $ser = New-Object Web.Script.Serialization.JavaScriptSerializer
  $links = @(); $n = 0
  foreach ($p in $plan) { $n++; $links += ('[' + $ser.Serialize([string]$p.title) + ', ' + $n + ']') }
  $linksJs = '[' + ($links -join ', ') + ']'
  $blankHtml = [IO.File]::ReadAllText((Join-Path $skill 'assets\page-blank.html'), [Text.Encoding]::UTF8)
  $jobPages = @(); $n = 0; $made = @(); $kept = @()
  foreach ($p in $plan) {
    $n++
    $file = '{0:D2}-{1}.html' -f $n, $p.id
    $dest = Join-Path $pagesDir $file
    $jobPages += @{ html = $file }
    if ((Test-Path -LiteralPath $dest) -and -not $Force) { $kept += $file; continue }
    if ($p.id -eq 'blank') {
      $safe = [Net.WebUtility]::HtmlEncode([string]$p.title)
      $html = $blankHtml.Replace('{{TITLE}}', $safe).Replace('{{LABEL}}', [Net.WebUtility]::HtmlEncode([string]$Name))
      if ($NoNav) { $html = $html.Replace('top: 96px; bottom: 48px;', 'top: 56px; bottom: 48px;') }
    } else {
      $html = [IO.File]::ReadAllText((Join-Path $skill ("assets\templates\$($p.id).html")), [Text.Encoding]::UTF8)
    }
    $html = $html.Replace('../kit/theme-cobalt.css', "../kit/theme-$themeName.css")
    if ($NoNav) { $html = [regex]::Replace($html, 'PH\.nav\(\{[^\r\n]*\}\);\r?\n?', '') }
    else {
      $nav = 'PH.nav({ brand: ' + $ser.Serialize([string]$Name) + ', links: ' + $linksJs + ', page: ' + $n + ' });'
      $html = [regex]::Replace($html, 'PH\.nav\(\{[^\r\n]*\}\);', { param($m) $nav })
    }
    # jumps that point past the last slide are pulled back to it
    $count = $plan.Count
    $html = [regex]::Replace($html, 'data-goto="(\d+)"', { param($m) 'data-goto="' + [Math]::Min([int]$m.Groups[1].Value, $count) + '"' })
    [IO.File]::WriteAllText($dest, $html, (New-Object Text.UTF8Encoding($false)))
    $made += $file
  }
  $jobFile = Join-Path $root 'job.json'
  if (-not (Test-Path -LiteralPath $jobFile) -or $Force) {
    $job = [ordered]@{ out = "$Name.pptm"; htmlDir = 'pages'; pages = @($jobPages); share = $true }
    Write-JsonFile $jobFile $job
  }
  Say "Deck project: $root"
  Say "  kit\            theme '$themeName', kit.css, kit.js, kit-deck.js (a toolbox: use what helps, leave the rest)"
  if ($designOut) { Say "  design\         design.md (the look to keep to), reference pictures, style.json" }
  foreach ($f in $made) { Say "  pages\$f" }
  foreach ($f in $kept) { Say "  pages\$f   (already there, kept; -Force overwrites)" }
  Say "  job.json        build settings"
  if ($skipped.Count -gt 0) { Say ("  NOTE: not copied, because this PC does not let PowerShell read them: " + ($skipped -join ', ')) }
  Say ''
  if ($designOut) { Say 'Next: read design\design.md and look at its pictures, then write each page so that it keeps to that look.' }
  else { Say 'Next: write each page (an empty page is a free canvas; an example page says at its bottom what to change).' }
  Say 'Then build:'
  Say ('  powershell -NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $PSScriptRoot 'html2ppt.ps1') + '" -Job "' + $jobFile + '"')
} catch {
  if ($exit -eq 0) { $exit = 1 }
  Say "FAILED: $($_.Exception.Message)"
  if ($exit -eq 1) { Say $_.ScriptStackTrace }
}
exit $exit
