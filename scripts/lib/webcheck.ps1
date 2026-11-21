# Dot-source (after common.ps1 and edge.ps1). Opens a finished web version in Edge without a window, walks through
# its slides and reports what a viewer would get: script errors, live pages that did not start, pictures.
# Plain ASCII on purpose (see hd_core.ps1).

# Returns @{ checked; slides; live; liveOk; errors = @(); shots = @(); why }.
function Test-WebDeck {
  param([Parameter(Mandatory)][string]$File, [Parameter(Mandatory)][string]$WorkDir, [string]$ShotsDir = '', [int]$SettleMs = 1200, [int]$MaxShots = 12)
  $res = @{ checked = $false; slides = 0; live = 0; liveOk = 0; errors = @(); shots = @(); why = '' }
  $edge = $null
  try {
    $html = [IO.File]::ReadAllText($File, [Text.Encoding]::UTF8)
    $edge = Start-EdgeTool $WorkDir -NoTools
    if (-not $edge) { $res.why = Get-EdgeWhyNot; return $res }
    $edge.SetViewport(1280, 764, 1.0)
    if (-not $edge.SetContent($html, 90000)) { $res.why = 'the web version did not finish loading in Edge'; return $res }
    $info = ConvertFrom-BigJson ([string]$edge.Eval('(function(){var v=[],k;for(k=0;k<DECK.slides.length;k++){if(!DECK.slides[k].hidden){v.push(DECK.slides[k]);}}if(!v.length){v=DECK.slides;}var live=[];for(k=0;k<v.length;k++){if(v[k].page){live.push(k+1);}}return JSON.stringify({slides:v.length,live:live,w:Math.round(DECK.slideWidth*4/3),h:Math.round(DECK.slideHeight*4/3)});})()', $false, 20000))
    $res.slides = [int]$info['slides']; $res.live = @($info['live']).Count
    $edge.SetViewport([int]$info['w'], [int]$info['h'] + 44, 1.0)
    $n = 0
    foreach ($pos in @($info['live'])) {
      [void]$edge.Eval("location.hash = '#$pos'; 'ok'", $false, 10000)
      Start-Sleep -Milliseconds ([Math]::Max(300, $SettleMs))
      $state = [string]$edge.Eval('(function(){var f=document.querySelectorAll("#stage iframe"),k,x=null;for(k=0;k<f.length;k++){if(f[k].style.display!=="none"){x=f[k];}}if(!x){return "noframe";}try{var d=x.contentDocument;return (d.body&&d.body.getAttribute("data-ih-ready")==="1"?"ready":"notready")+"|"+(d.documentElement.getAttribute("data-ih-err")||"");}catch(e){return "blocked|"+e.message;}})()', $false, 20000)
      $part = $state -split '\|', 2
      if ($part[0] -eq 'ready' -and -not $part[1]) { $res.liveOk++ }
      else { $res.errors += ("slide {0}: {1}{2}" -f $pos, $part[0], $(if ($part.Count -gt 1 -and $part[1]) { ' - ' + $part[1] } else { '' })) }
      if ($ShotsDir -and $n -lt $MaxShots) {
        $shot = Join-Path $ShotsDir ("web_slide_{0:D2}.png" -f [int]$pos)
        try { $edge.Screenshot($shot); $res.shots += $shot } catch {}
      }
      $n++
    }
    if ($res.live -eq 0 -and $ShotsDir) {
      $shot = Join-Path $ShotsDir 'web_slide_01.png'
      try { Start-Sleep -Milliseconds 400; $edge.Screenshot($shot); $res.shots += $shot } catch {}
    }
    foreach ($e in $edge.Errors) { if ($res.errors -notcontains [string]$e) { $res.errors += [string]$e } }
    $res.checked = $true
  } catch {
    $x = $_.Exception; while ($x.InnerException) { $x = $x.InnerException }
    $res.why = $x.Message
  } finally { Stop-EdgeTool }
  return $res
}
