# Dot-source (after common.ps1). Microsoft Edge without a window as a build tool: converts modern CSS / JavaScript
# to what the IE11 engine understands, cuts fonts down, draws reference pictures, checks the web version.
# Edge is part of Windows 10 / 11. When it cannot be used, the callers carry on without these steps and say so.
# Plain ASCII on purpose (see hd_core.ps1).

$script:EdgeSession = $null
$script:EdgeWhyNot = ''
$script:EdgeFontTool = $false

function Get-EdgeWhyNot { return $script:EdgeWhyNot }

# Starts the tool browser (once per run). Returns the session, or $null when Edge cannot be used (Get-EdgeWhyNot says why).
function Start-EdgeTool([Parameter(Mandatory)][string]$WorkDir, [switch]$NoTools) {
  if ($script:EdgeSession) { return $script:EdgeSession }
  if ($script:EdgeWhyNot) { return $null }
  # maintenance hook: PPTHTML_NO_EDGE=1 behaves like a PC without a usable Edge
  if ($env:PPTHTML_NO_EDGE -eq '1') { $script:EdgeWhyNot = 'switched off by PPTHTML_NO_EDGE'; return $null }
  try {
    if (-not ('EdgeKit.Session' -as [type])) {
      Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $script:HdScripts 'cs\EdgeKit.cs'))) -ReferencedAssemblies System, System.Core, System.Web.Extensions
    }
    $exe = [EdgeKit.Session]::FindBrowser()
    if (-not $exe) { $script:EdgeWhyNot = 'Microsoft Edge (or Chrome) was not found on this PC'; return $null }
    $tools = Join-Path $script:HdSkill 'assets\tools'
    foreach ($f in 'babel.min.js', 'convert.js', 'tool.js') {
      if (-not (Test-Path -LiteralPath (Join-Path $tools $f))) { $script:EdgeWhyNot = "the tool file assets\tools\$f is missing from the skill folder"; return $null }
    }
    $prof = Join-Path $WorkDir ('edge_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    $s = New-Object EdgeKit.Session($exe, $prof, 1280, 720, 25000)
    try {
      if ($NoTools) { $script:EdgeSession = $s; return $s }
      $s.LoadScript((Join-Path $tools 'babel.min.js'), 60000)
      $s.LoadScript((Join-Path $tools 'convert.js'), 60000)
      $s.LoadScript((Join-Path $tools 'tool.js'), 20000)
      if ($s.Eval('typeof IHTool + typeof IHConvert + typeof Babel', $false, 10000) -ne 'objectobjectobject') { throw 'the tool scripts did not load' }
    } catch { try { $s.Dispose() } catch {}; throw }
    $script:EdgeSession = $s
    return $s
  } catch {
    $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
    $script:EdgeWhyNot = "Edge could not be used as a build tool: $($e.Message)"
    return $null
  }
}

function Stop-EdgeTool {
  if ($script:EdgeSession) { try { $script:EdgeSession.Dispose() } catch {}; $script:EdgeSession = $null }
  $script:EdgeFontTool = $false
}

# True while the tool browser still answers. Its connection can be closed in the middle of a run (seen once on the
# measured PC: "An established connection was aborted by the software in your host machine").
function Test-EdgeTool {
  if (-not $script:EdgeSession) { return $false }
  try { return ([string]$script:EdgeSession.Eval('"ok"', $false, 5000) -eq 'ok') } catch { return $false }
}

# Starts the tool browser again after it went away (at most twice per run). Returns the new session, or $null.
$script:EdgeRestarts = 0
function Restart-EdgeTool([Parameter(Mandatory)][string]$WorkDir, [switch]$NoTools) {
  if ($script:EdgeRestarts -ge 2) { return $null }
  $script:EdgeRestarts++
  Stop-EdgeTool
  Start-Sleep -Milliseconds 800
  if ($NoTools) { return (Start-EdgeTool $WorkDir -NoTools) }
  return (Start-EdgeTool $WorkDir)
}

# JSON text for a .NET object (strings of any size), without the 2 MB limit of ConvertTo-Json's serializer defaults.
function ConvertTo-BigJson($Object) {
  Add-Type -AssemblyName System.Web.Extensions
  $ser = New-Object Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = [int]::MaxValue; $ser.RecursionLimit = 400
  return $ser.Serialize($Object)
}
function ConvertFrom-BigJson([string]$Text) {
  Add-Type -AssemblyName System.Web.Extensions
  $ser = New-Object Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = [int]::MaxValue; $ser.RecursionLimit = 400
  return $ser.DeserializeObject($Text)
}

# Calls a function of the tool page with string arguments; returns its (string) result. Promises are awaited.
function Invoke-EdgeTool($Session, [string]$Function, [object[]]$Arguments = @(), [int]$TimeoutMs = 120000) {
  $names = @()
  for ($i = 0; $i -lt $Arguments.Count; $i++) {
    $a = $Arguments[$i]
    if ($a -is [string]) { $n = "__a$i"; $Session.SetText($n, $a, 60000); $names += "window.$n" }
    elseif ($a -is [bool]) { $names += $(if ($a) { 'true' } else { 'false' }) }
    else { $names += ([string]$a) }
  }
  return $Session.Eval("Promise.resolve($Function($($names -join ', ')))", $true, $TimeoutMs)
}

# Converts one prepared page (PptHtml.PrepState) in place. Returns $true when the conversion was applied.
function Convert-PageForIe($Session, $State, [int]$CanvasW, [int]$CanvasH) {
  $job = New-Object 'Collections.Generic.Dictionary[string,object]'
  $job['canvasW'] = $CanvasW; $job['canvasH'] = $CanvasH
  $scripts = New-Object Collections.ArrayList
  foreach ($code in $State.Scripts) { $d = New-Object 'Collections.Generic.Dictionary[string,object]'; $d['code'] = $code; [void]$scripts.Add($d) }
  $job['scripts'] = $scripts.ToArray()
  $job['styles'] = [string[]]$State.Styles
  $job['inline'] = [string[]]$State.InlineStyles
  $resText = Invoke-EdgeTool $Session 'IHTool.convert' @((ConvertTo-BigJson $job)) 180000
  $res = ConvertFrom-BigJson $resText
  $outScripts = New-Object string[] ($State.Scripts.Length)
  for ($i = 0; $i -lt $outScripts.Length; $i++) { $v = $res['scripts'][$i]; if ($v -ne $null) { $outScripts[$i] = [string]$v } }
  $inS = New-Object string[] ($State.InlineStyles.Length); $inV = New-Object string[] ($State.InlineStyles.Length)
  for ($i = 0; $i -lt $inS.Length; $i++) {
    $v = $res['inline'][$i]
    if ($v -ne $null) { $inS[$i] = [string]$v['style']; if ($v['v'] -ne $null) { $inV[$i] = ConvertTo-BigJson $v['v'] } }
  }
  $css = $null; if ($res['css'] -ne $null) { $css = [string]$res['css'] }
  $table = ConvertTo-BigJson $res['table']
  $notes = @(); foreach ($x in $res['notes']) { $notes += [string]$x }
  foreach ($x in $res['warnings']) { $notes += ('LOOK: ' + [string]$x) }
  $allow = @(); foreach ($x in $res['allow']) { $allow += [string]$x }
  [PptHtml.HtmlPrep]::ApplyConversion($State, $outScripts, $css, $inS, $inV, $table, [string[]]$allow, [string[]]$notes)
  return $true
}
