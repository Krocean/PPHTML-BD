# Dot-source (after common.ps1). The web version of a presentation: one self-contained HTML file that opens in
# any current browser. Slides are pictures; where the presentation carries live HTML pages (built by this
# skill), the pages themselves run in place. Also: the share bundle (folder + zip with a short read-me).

function ConvertTo-HtmlText([string]$s) { return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;') }

# Source: .pptx / .pptm (any path PowerPoint can open). Needs the hidden PowerPoint (Start-Hd) to be running.
# Returns @{ out; warnings; slides; livePages; bytes }.
function New-WebDeck {
  param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$OutFile, [Parameter(Mandatory)][string]$WorkDir, [string]$Title = '', [int]$Width = 1600, [string]$ImageFormat = 'png')
  $outDir = Join-Path $WorkDir 'web'
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force $outDir | Out-Null }
  $run = Invoke-HdChild -Script 'deck_extract.ps1' -WorkDir $WorkDir -Params @{ file = $Source; outDir = $outDir; width = $Width; imageFormat = $ImageFormat } -TimeoutSec 900
  if (-not $run.Ok) { throw (Format-ChildFailure $run 'Reading the slides') }
  $x = $run.Result
  $warnings = @()
  Initialize-HtmlPrep

  $pages = @()
  if ($x.pagesXml) {
    $doc = New-Object Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.Load([string]$x.pagesXml)
    $headH = Read-Asset 'harness-head.txt'; $tailWeb = Read-Asset 'harness-tail-web.txt'
    foreach ($pg in $doc.DocumentElement.ChildNodes) {
      if ($pg.LocalName -ne 'page') { continue }
      $html = ''; $htmlM = ''; $bg = 16777215; $w = 0; $h = 0
      foreach ($c in $pg.ChildNodes) {
        switch ($c.LocalName) { 'html' { $html = $c.InnerText } 'htmlm' { $htmlM = $c.InnerText } 'bg' { $bg = [int]$c.InnerText } 'w' { $w = [int]$c.InnerText } 'h' { $h = [int]$c.InnerText } }
      }
      # a browser gets the page as it was written (the document of the modern engine) where the file carries it
      if ($htmlM) { $html = $htmlM }
      $pages += [pscustomobject]@{ w = $w; h = $h; bg = ('#{0:X2}{1:X2}{2:X2}' -f ($bg -band 255), (($bg -shr 8) -band 255), (($bg -shr 16) -band 255)); html = [PptHtml.HtmlPrep]::SwapHarness($html, $headH, $tailWeb) }
    }
  }

  $mime = if ($ImageFormat -eq 'jpg') { 'image/jpeg' } else { 'image/png' }
  $slides = @(); $live = 0
  foreach ($s in @($x.slides)) {
    $links = @()
    foreach ($l in @($s.links)) { if ($l) { $links += [ordered]@{ l = [Math]::Round([double]$l.left, 1); t = [Math]::Round([double]$l.top, 1); w = [Math]::Round([double]$l.width, 1); h = [Math]::Round([double]$l.height, 1); href = [string]$l.href; slide = [int]$l.slide } } }
    $page = [int]$s.page
    if ($page -gt $pages.Count) { $page = 0 }
    # what the page stored in PowerPoint goes along (plain URI-encoded text; anything else is left out)
    $st = [string]$s.state
    if ($page -gt 0 -and $st -and $st -match '^[A-Za-z0-9%\-_.!~*''()]+$') {
      $ph = [string]$pages[$page - 1].html; $mark = [PptHtml.HtmlPrep]::HeadClose; $at = $ph.IndexOf($mark)
      if ($at -ge 0) { $pages[$page - 1].html = $ph.Substring(0, $at + $mark.Length) + '<script>window.__ihState="' + $st + '";</script>' + $ph.Substring($at + $mark.Length) }
    }
    $area = $null
    if ($page -gt 0 -and $s.area) { $area = [ordered]@{ l = [double]$s.area.left; t = [double]$s.area.top; w = [double]$s.area.width; h = [double]$s.area.height }; $live++ }
    $slides += [ordered]@{ n = [int]$s.index; title = [string]$s.title; hidden = [bool]$s.hidden; image = "@@IMG$($s.index)@@"; text = [string]$s.text; notes = [string]$s.notes; links = @($links); page = $page; area = $area }
  }
  $pg2 = @(); $k = 0
  foreach ($p in $pages) { $k++; $pg2 += [ordered]@{ w = $p.w; h = $p.h; bg = $p.bg; html = "@@PAGE$k@@" } }
  if (-not $Title) { $Title = [IO.Path]::GetFileNameWithoutExtension([string]$x.name) }
  $meta = [ordered]@{ title = $Title; slideWidth = [double]$x.slideWidth; slideHeight = [double]$x.slideHeight; slides = @($slides); pages = @($pg2) }
  $json = $meta | ConvertTo-Json -Depth 8 -Compress

  $tpl = [IO.File]::ReadAllText((Join-Path $script:HdSkill 'assets\viewer-template.txt'), [Text.Encoding]::UTF8)
  $tpl = $tpl.Replace('__TITLE__', (ConvertTo-HtmlText $Title))
  $cut = $tpl.IndexOf('/*__DECK__*/null')
  if ($cut -lt 0) { throw 'viewer template: data marker missing' }
  $tmp = Join-Path $outDir 'deck.html'
  $sw = New-Object IO.StreamWriter($tmp, $false, $script:Utf8NoBom)
  try {
    $sw.Write($tpl.Substring(0, $cut))
    $pos = 0
    foreach ($m in [regex]::Matches($json, '"@@(IMG|PAGE)(\d+)@@"')) {
      $sw.Write($json.Substring($pos, $m.Index - $pos))
      $num = [int]$m.Groups[2].Value
      if ($m.Groups[1].Value -eq 'IMG') {
        $file = Join-Path $outDir ([string](@($x.slides) | Where-Object { [int]$_.index -eq $num } | Select-Object -First 1).image)
        $sw.Write('"data:' + $mime + ';base64,')
        $sw.Write([Convert]::ToBase64String([IO.File]::ReadAllBytes($file)))
        $sw.Write('"')
      } else {
        $sw.Write([PptHtml.HtmlPrep]::JsString([string]$pages[$num - 1].html))
      }
      $pos = $m.Index + $m.Length
    }
    $sw.Write($json.Substring($pos))
    $sw.Write($tpl.Substring($cut + '/*__DECK__*/null'.Length))
  } finally { $sw.Dispose() }
  $bytes = (Get-Item -LiteralPath $tmp).Length
  if ($bytes -gt 40MB) { $warnings += "The web version is $([Math]::Round($bytes / 1MB)) MB. For a smaller file use -Jpg and/or a smaller -Width." }
  $written = Copy-Deliver $tmp $OutFile
  if (-not $written) { $warnings += "Could not write $OutFile. The web version is here: $tmp"; $written = $tmp }
  return @{ out = $written; warnings = $warnings; slides = $slides.Count; livePages = $live; bytes = $bytes }
}

# A folder (and a zip of it) to hand to other people: the .pptm, the web version and a read-me that says how
# to allow the macros of a received file and what to use when that is not possible.
function New-ShareBundle {
  param([Parameter(Mandatory)][string]$Pptm, [Parameter(Mandatory)][string]$Web, [Parameter(Mandatory)][string]$WorkDir, [Parameter(Mandatory)]$Strings)
  Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
  $name = [IO.Path]::GetFileNameWithoutExtension($Pptm) + [string]$Strings.shareSuffix
  $stageRoot = Join-Path $WorkDir 'share'
  $stage = Join-Path $stageRoot $name
  if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force }
  New-Item -ItemType Directory -Force $stage | Out-Null
  Invoke-Retry { [IO.File]::Copy($Pptm, (Join-Path $stage ([IO.Path]::GetFileName($Pptm))), $true) } | Out-Null
  [IO.File]::Copy($Web, (Join-Path $stage ([IO.Path]::GetFileName($Web))), $true)
  $readme = (@($Strings.readme) -join "`r`n").Replace('{pptm}', [IO.Path]::GetFileName($Pptm)).Replace('{html}', [IO.Path]::GetFileName($Web)) + "`r`n"
  [IO.File]::WriteAllText((Join-Path $stage ([string]$Strings.readmeName)), $readme, (New-Object Text.UTF8Encoding($true)))
  # The zip is written entry by entry with "/" in the names: the one-call method of this .NET version writes
  # "\" there, which other systems (macOS, Linux) do not read as folders.
  $zipTmp = Join-Path $stageRoot ($name + '.zip')
  $zs = [IO.File]::Open($zipTmp, 'Create')
  $za = New-Object IO.Compression.ZipArchive($zs, [IO.Compression.ZipArchiveMode]::Create, $false, [Text.Encoding]::UTF8)
  try {
    foreach ($f in Get-ChildItem -LiteralPath $stage -File) {
      [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($za, $f.FullName, ($name + '/' + $f.Name), [IO.Compression.CompressionLevel]::Optimal)
    }
  } finally { $za.Dispose(); $zs.Dispose() }
  $destDir = Split-Path -Parent $Pptm
  $folder = Join-Path $destDir $name
  $okFolder = $true
  try {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Force $folder | Out-Null }
    foreach ($f in Get-ChildItem -LiteralPath $stage -File) { [IO.File]::Copy($f.FullName, (Join-Path $folder $f.Name), $true) }
  } catch { $okFolder = $false; $folder = $stage }
  $zip = Copy-Deliver $zipTmp (Join-Path $destDir ($name + '.zip'))
  if (-not $zip) { $zip = $zipTmp }
  return @{ folder = $folder; zip = $zip; delivered = $okFolder }
}
