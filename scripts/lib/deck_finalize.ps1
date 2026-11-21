param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): opens the staged presentation (now with the macro project inside), gives
# every still its fallback action (run macro InlineHtml_Activate on click and on mouse over) and saves the
# final .pptm. Opening and saving it once in PowerPoint also stores the project the way PowerPoint itself does.
#   args: file (staged .pptm), saveTo
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
$pres = $null; $app = $null; $oldSec = $null
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 1 } catch {}
  $pres = $app.Presentations.Open([string]$A.file, 0, 0, 0)
  $hasVba = $false; try { $hasVba = [bool]$pres.HasVBProject } catch {}
  $done = 0
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $tag = ''; try { $tag = [string]$sl.Tags.Item('HtmlPage') } catch {}
    if (-not $tag) { continue }
    $sh = $null; try { $sh = $sl.Shapes.Item('htmlArea') } catch {}
    if (-not $sh) { L "slide ${i}: tagged but no htmlArea shape"; continue }
    foreach ($k in 1, 2) {
      $as = $sh.ActionSettings.Item($k)
      $as.Action = 8
      $as.Run = 'InlineHtml_Activate'
    }
    $done++
  }
  if (Test-Path -LiteralPath ([string]$A.saveTo)) { Remove-Item -LiteralPath ([string]$A.saveTo) -Force }
  $pres.SaveAs([string]$A.saveTo, 25)
  L "actions set on $done slide(s); hasVBProject=$hasVba; saved $($A.saveTo)"
  Write-JsonFile $A.result ([ordered]@{ slidesWithHtml = $done; hasVba = $hasVba })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
