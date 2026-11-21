# Dot-source (after common.ps1 and edge.ps1). Font embedding: the fonts a page is set in are cut down to the
# characters of that page and put inside it (WOFF, data URI), so the page looks the same on a PC that does not
# have them. Only fonts whose own licence flags allow embedding are taken; the rest is left to the PC.
# Sources, in this order: fonts installed on this PC, then the skill's font pack (assets\fonts).
# Plain ASCII on purpose (see hd_core.ps1).

$script:FontIndex = $null
$script:FontLoaded = @{}
$script:GenericFamilies = @('serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'math', 'emoji', 'fangsong', '-apple-system', 'blinkmacsystemfont', 'inherit', 'initial', 'unset')

function Get-FontIndex {
  if ($script:FontIndex) { return $script:FontIndex }
  if (-not ('FontKit.Fonts' -as [type])) { Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\FontKit.cs'))) }
  $idx = @{}
  $dirs = @(@((Join-Path $env:WINDIR 'Fonts'), $false), @((Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'), $false), @((Join-Path $script:HdSkill 'assets\fonts'), $true))
  # maintenance hook: PPTHTML_NO_SYSTEM_FONTS=1 behaves like a PC that has only the skill's font pack
  if ($env:PPTHTML_NO_SYSTEM_FONTS -eq '1') { $dirs = @(, $dirs[2]) }
  foreach ($d in $dirs) {
    foreach ($f in [FontKit.Fonts]::Scan([string]$d[0], [bool]$d[1])) {
      foreach ($n in $f.Families) {
        $k = $n.ToLowerInvariant()
        if (-not $idx.ContainsKey($k)) { $idx[$k] = New-Object Collections.ArrayList }
        [void]$idx[$k].Add($f)
      }
    }
  }
  $script:FontIndex = $idx
  return $idx
}

$script:HangulCache = @{}
function Test-FaceHangul($Face) {
  $key = '{0}#{1}' -f $Face.File, $Face.Index
  if (-not $script:HangulCache.ContainsKey($key)) { $script:HangulCache[$key] = [FontKit.Fonts]::HasChar([string]$Face.File, [int]$Face.Index, 0xAC00) }
  return $script:HangulCache[$key]
}

# The fonts of a CSS font stack that this PC would draw Korean + Latin text with, each with the face for the given
# weight: every family from the front of the stack up to the first one that has Hangul (a Latin font is followed by a
# Korean one). Returns a list of { family, skip } (cannot be embedded) / { family, face, pin, cssWeight }.
function Resolve-FontFace([string]$Stack, [int]$Weight) {
  $idx = Get-FontIndex
  $out = @()
  foreach ($raw in ($Stack -split ',')) {
    $name = $raw.Trim().Trim('"').Trim("'").Trim()
    if (-not $name) { continue }
    $k = $name.ToLowerInvariant()
    if ($script:GenericFamilies -contains $k) { break }
    if (-not $idx.ContainsKey($k)) { continue }
    $faces = @($idx[$k])
    $sys = @($faces | Where-Object { -not $_.Bundled })
    if ($sys.Count -gt 0) { $faces = $sys }
    $ok = @($faces | Where-Object { $_.Embeddable })
    if ($ok.Count -eq 0) {
      $out += [pscustomobject]@{ family = $name; skip = 'licence'; file = [string]$faces[0].File }
      if (Test-FaceHangul $faces[0]) { break } else { continue }
    }
    $upright = @($ok | Where-Object { -not $_.Italic })
    if ($upright.Count -gt 0) { $ok = $upright }
    $var = @($ok | Where-Object { $_.Variable }) | Select-Object -First 1
    if ($var) {
      $w = [Math]::Max([int]$var.WghtMin, [Math]::Min([int]$var.WghtMax, $Weight))
      $out += [pscustomobject]@{ family = $name; skip = ''; face = $var; pin = $w; cssWeight = $Weight }
      if (Test-FaceHangul $var) { break } else { continue }
    }
    $best = $ok | Sort-Object { [Math]::Abs([int]$_.Weight - $Weight) } | Select-Object -First 1
    $out += [pscustomobject]@{ family = $name; skip = ''; face = $best; pin = 0; cssWeight = [int]$best.Weight }
    if (Test-FaceHangul $best) { break }
  }
  return $out
}

# Embeds the fonts of the page that is currently shown in the tool browser (IHTool.show) into its PrepState.
# Returns notes for the report.
function Add-PageFonts($Session, $State, [int]$MaxFaces = 12) {
  $notes = @()
  if (-not $script:EdgeFontTool) {
    $Session.LoadScript((Join-Path $script:HdSkill 'assets\tools\fonttool.js'), 60000)
    [void]$Session.Eval('IHFont.init()', $true, 60000)
    $script:EdgeFontTool = $true
    $script:FontLoaded = @{}
  }
  $uses = ConvertFrom-BigJson ([string](Invoke-EdgeTool $Session 'IHTool.fontUse' @() 30000))
  $want = [ordered]@{}; $licence = @()
  foreach ($u in $uses) {
    foreach ($r in @(Resolve-FontFace ([string]$u['stack']) ([int]$u['weight']))) {
      if (-not $r) { continue }
      if ($r.skip) { if ($r.file -notlike (Join-Path $env:WINDIR 'Fonts\*') -and $licence -notcontains $r.family) { $licence += $r.family }; continue }
      $key = '{0}|{1}|{2}|{3}' -f $r.family.ToLowerInvariant(), $r.face.File, $r.face.Index, $r.pin
      if (-not $want.Contains($key)) { $want[$key] = $r }
    }
  }
  if ($licence.Count -gt 0) { $notes += "fonts left to the PC (their licence flags do not allow embedding): $($licence -join ', ')" }
  if ($want.Count -eq 0) { return $notes }
  $list = @($want.Values)
  if ($list.Count -gt $MaxFaces) { $notes += "the page uses $($list.Count) font faces; only the first $MaxFaces were embedded"; $list = @($list | Select-Object -First $MaxFaces) }
  $chars = [FontKit.Fonts]::Characters([string]$State.Doc)
  $css = New-Object Text.StringBuilder
  $done = @(); $bytes = 0
  foreach ($r in $list) {
    $fkey = '{0}#{1}' -f $r.face.File, $r.face.Index
    try {
      if (-not $script:FontLoaded.ContainsKey($r.face.File)) {
        $len = (Get-Item -LiteralPath $r.face.File).Length
        if ($len -gt 60MB) { $notes += "font file too large to embed: $([IO.Path]::GetFileName($r.face.File))"; continue }
        $Session.SetText('__font', [Convert]::ToBase64String([IO.File]::ReadAllBytes($r.face.File)), 120000)
        [void]$Session.Eval("IHFont.add($(ConvertTo-BigJson ([string]$r.face.File)), window.__font); window.__font = null; 'ok'", $false, 120000)
        $script:FontLoaded[$r.face.File] = $true
      }
      $Session.SetText('__chars', $chars, 30000)
      $b64 = $Session.Eval("IHFont.subset($(ConvertTo-BigJson ([string]$r.face.File)), window.__chars, $([int]$r.pin), $([int]$r.face.Index))", $false, 120000)
      if (-not $b64) { throw 'empty result' }
      [void]$css.Append("@font-face{font-family:`"$($r.family)`";font-style:normal;font-weight:$($r.cssWeight);src:url(data:font/woff;base64,$b64) format(`"woff`");}`n")
      $bytes += [int]($b64.Length * 3 / 4)
      $done += ('{0} {1}{2}' -f $r.family, $r.cssWeight, $(if ($r.face.Bundled) { ' (font pack)' } else { '' }))
    } catch {
      # the tool browser itself went away: let the caller start it again and repeat the page (nothing was added to the page yet)
      if (-not (Test-EdgeTool)) { throw }
      $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
      $notes += "font $($r.family) $($r.cssWeight) could not be embedded: $($e.Message)"
    }
  }
  if ($done.Count -gt 0) {
    [PptHtml.HtmlPrep]::AddCss($State, $css.ToString())
    $notes += "fonts embedded for other PCs ($([Math]::Round($bytes / 1KB)) KB): $($done -join ', ')"
  }
  return $notes
}
