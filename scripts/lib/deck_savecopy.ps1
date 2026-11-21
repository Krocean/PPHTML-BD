param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): saves a copy of a presentation somewhere. Used when PowerShell itself is
# not allowed to write to the user's folder (protected folders) but PowerPoint is.
#   args: file, saveTo
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
$pres = $null; $app = $null; $oldSec = $null
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}
  $pres = $app.Presentations.Open([string]$A.file, -1, 0, 0)
  $pres.SaveCopyAs([string]$A.saveTo)
  Write-ChildLog $A.log "copy saved to $($A.saveTo)"
  Write-JsonFile $A.result ([ordered]@{ saved = [string]$A.saveTo })
} catch {
  Write-ChildLog $A.log "ERROR: $($_.Exception.Message)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
