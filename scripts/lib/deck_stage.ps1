param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (hidden PowerPoint): opens the user's presentation read-only, saves a working copy and reports
# what is in it: slide size, slides, places marked for HTML, whether it carries macros.
#   args: deck (original path), copyTo (working copy, same extension)
# A place for HTML is a shape whose text (or alternative text) is  {{html:file.html}}  or a shape named htmlArea
# (left by an earlier run of this tool).
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
$pres = $null; $app = $null; $oldSec = $null
$markerRx = '^\s*(?:\{\{|\[\[)\s*html\s*:\s*(.+?)\s*(?:\}\}|\]\])\s*$'
try {
  $app = Get-HdApp
  Close-HdExtra $app
  try { $oldSec = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}
  $pres = $app.Presentations.Open([string]$A.deck, -1, 0, 0)
  $pres.SaveCopyAs([string]$A.copyTo)
  $deckTag = ''; try { $deckTag = [string]$pres.Tags.Item('InlineHtml') } catch {}
  $hasVba = $false; try { $hasVba = [bool]$pres.HasVBProject } catch {}
  $slides = @()
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $title = ''
    try { if ($sl.Shapes.HasTitle) { $title = [string]$sl.Shapes.Title.TextFrame.TextRange.Text } } catch {}
    $pageTag = ''; try { $pageTag = [string]$sl.Tags.Item('HtmlPage') } catch {}
    $areas = @()
    for ($k = 1; $k -le $sl.Shapes.Count; $k++) {
      $sh = $sl.Shapes.Item($k)
      $kind = ''; $file = ''
      $txt = ''
      try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = [string]$sh.TextFrame.TextRange.Text } } catch {}
      if ($txt -match $markerRx) { $kind = 'marker'; $file = $Matches[1] }
      if (-not $kind) {
        $alt = ''; try { $alt = [string]$sh.AlternativeText } catch {}
        if ($alt -match $markerRx) { $kind = 'marker'; $file = $Matches[1] }
      }
      if (-not $kind -and $sh.Name -eq 'htmlArea') { $kind = 'area' }
      if ($kind) {
        $areas += [pscustomobject][ordered]@{ kind = $kind; file = $file; id = [int]$sh.Id; name = [string]$sh.Name; left = [double]$sh.Left; top = [double]$sh.Top; width = [double]$sh.Width; height = [double]$sh.Height; z = [int]$sh.ZOrderPosition }
      }
    }
    $hidden = $false; try { $hidden = ($sl.SlideShowTransition.Hidden -ne 0) } catch {}
    $slides += [pscustomobject][ordered]@{ index = $i; title = $title.Trim(); hidden = $hidden; pageTag = $pageTag; areas = @($areas) }
  }
  $res = [ordered]@{
    name = [string]$pres.Name
    slideWidth = [double]$pres.PageSetup.SlideWidth
    slideHeight = [double]$pres.PageSetup.SlideHeight
    slideCount = [int]$pres.Slides.Count
    hasVba = $hasVba
    inlineDeck = ($deckTag -eq '1')
    version = [string]$app.Version
    slides = @($slides)
  }
  L "opened [$($pres.Name)] slides=$($pres.Slides.Count) size=$($res.slideWidth)x$($res.slideHeight) hasVba=$hasVba inlineDeck=$($res.inlineDeck)"
  Write-JsonFile $A.result $res
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
} finally {
  try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
  try { if ($app -and $oldSec -ne $null) { $app.AutomationSecurity = $oldSec } } catch {}
}
