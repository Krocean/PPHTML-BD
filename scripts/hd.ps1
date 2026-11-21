param([Parameter(Position = 0)][ValidateSet('status', 'start', 'stop')][string]$Action = 'status')
# Maintenance: the hidden PowerPoint instance this skill works with.
#   hd.ps1 status   is it running?
#   hd.ps1 stop     end it (use after a run that was interrupted)
#   hd.ps1 start    start it (the main scripts do this themselves)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\common.ps1')
Initialize-Console
switch ($Action) {
  'start' { Say ("hidden PowerPoint: " + (Start-Hd)) }
  'stop' { Say ("hidden PowerPoint: " + (Stop-Hd)) }
  default {
    $p = Get-HdPid
    if (-not $p) { Say 'hidden PowerPoint: not running'; break }
    try {
      $app = Get-HdApp
      $names = @(); for ($i = 1; $i -le $app.Presentations.Count; $i++) { $names += $app.Presentations.Item($i).Name }
      Say "hidden PowerPoint: running (process $p, version $($app.Version), open: $($names -join ', '))"
    } catch { Say "hidden PowerPoint: process $p exists but does not answer; run 'hd.ps1 stop'" }
  }
}
