param(
  [string]$Out = '',     # folder to write into (default: the current folder)
  [switch]$NoFonts       # leave the font pack out (3.5 MB smaller; PCs without the Noto Korean fonts then fall back to system fonts)
)
# Packs the skill for handing to other people: ppt-html.zip (the skill folder, ready to unzip into
# %USERPROFILE%\.claude\skills or %USERPROFILE%\.agents\skills) and a short install note next to it. Nothing of this PC goes into the zip:
# only the files of the skill folder.
# Exit codes: 0 done; 1 error.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
Initialize-Console
$exit = 0
try {
  Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
  $skill = $script:HdSkill
  $strings = Read-JsonFile (Join-Path $skill 'assets\strings.json')
  $dest = if ($Out) { Resolve-UserPath $Out (Get-Location).Path } else { (Get-Location).Path }
  if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Force $dest | Out-Null }
  $work = New-WorkDir 'pack'
  $zipTmp = Join-Path $work 'ppt-html.zip'
  $skip = @('(^|\\)\.git(\\|$)', '^dev\\(node_modules|out)\\', '^dev\\package-lock\.json$', '\.(bak|tmp|log)$')
  if ($NoFonts) { $skip += '^assets\\fonts\\.*\.(ttf|otf)$' }
  $files = @(Get-ChildItem -LiteralPath $skill -Recurse -File | Where-Object {
    $rel = $_.FullName.Substring($skill.Length).TrimStart('\')
    $keep = $true
    foreach ($rx in $skip) { if ($rel -match $rx) { $keep = $false } }
    $keep
  })
  $zs = [IO.File]::Open($zipTmp, 'Create')
  $za = New-Object IO.Compression.ZipArchive($zs, [IO.Compression.ZipArchiveMode]::Create, $false, [Text.Encoding]::UTF8)
  $bytes = 0; $skipped = @()
  try {
    foreach ($f in $files) {
      $rel = $f.FullName.Substring($skill.Length).TrimStart('\').Replace('\', '/')
      try {
        $data = [IO.File]::ReadAllBytes($f.FullName)      # read first: an unreadable file must not leave an empty entry behind
        $entry = $za.CreateEntry(('ppt-html/' + $rel), [IO.Compression.CompressionLevel]::Optimal)
        $es = $entry.Open(); $es.Write($data, 0, $data.Length); $es.Dispose()
        $bytes += $f.Length
      } catch { $skipped += $rel }
    }
  } finally { $za.Dispose(); $zs.Dispose() }
  $zip = Copy-Deliver $zipTmp (Join-Path $dest 'ppt-html.zip')
  if (-not $zip) { $zip = $zipTmp }
  $noteTmp = Join-Path $work 'install.txt'
  $noteLines = @($strings.install)
  # files that could not be read on this PC are named in the note, so that the receiver knows what is not in the zip
  if ($skipped.Count -gt 0 -and $strings.installMissing) { $noteLines += ''; $noteLines += @($strings.installMissing); foreach ($x in $skipped) { $noteLines += "  - $x" } }
  [IO.File]::WriteAllText($noteTmp, (($noteLines -join "`r`n") + "`r`n"), (New-Object Text.UTF8Encoding($true)))
  $note = Copy-Deliver $noteTmp (Join-Path (Split-Path -Parent $zip) ([string]$strings.installName))
  Say ("{0}   ({1} files, {2:N1} MB unpacked, {3:N1} MB zipped)" -f $zip, ($files.Count - $skipped.Count), ($bytes / 1MB), ((Get-Item -LiteralPath $zip).Length / 1MB))
  if ($note) { Say $note }
  foreach ($x in $skipped) { Say "WARNING: not packed, because PowerShell may not read it on this PC (security software): $x" }
  Say ''
  Say 'To install on another PC: unzip so that %USERPROFILE%\.claude\skills\ppt-html\SKILL.md (Claude Code) or %USERPROFILE%\.agents\skills\ppt-html\SKILL.md (Codex, Gemini CLI) exists, then run scripts\doctor.ps1 there.'
} catch {
  $exit = 1
  Say "FAILED: $($_.Exception.Message)"
  Say $_.ScriptStackTrace
}
exit $exit
