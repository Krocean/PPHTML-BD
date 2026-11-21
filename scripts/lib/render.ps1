param([Parameter(Mandatory)][string]$ArgsFile)
# Child script (no PowerPoint needed): renders prepared pages in the Internet Explorer 11 engine, the same
# engine the slide show uses, off screen. For every page it saves a picture (the still that sits on the slide)
# and reports what the engine said: document mode, script errors, background colour, overflow.
#   args: pages = [{ n, file, canvasW, canvasH, png }], settleMs, scale, probe (JavaScript expression evaluated in every page; its value is reported)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'hd_core.ps1')
$A = Read-JsonFile $ArgsFile
function L([string]$s) { Write-ChildLog $A.log $s }
try {
  Add-Type -AssemblyName System.Windows.Forms
  $scale = 2; if ($A.scale) { $scale = [int]$A.scale }
  $settle = 1500; if ($A.settleMs -ne $null) { $settle = [int]$A.settleMs }
  function Pump([int]$ms) { $sw = [Diagnostics.Stopwatch]::StartNew(); while ($sw.ElapsedMilliseconds -lt $ms) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 15 } }
  $out = @()
  foreach ($p in @($A.pages)) {
    $w = [int]$p.canvasW * $scale; $h = [int]$p.canvasH * $scale
    $wb = New-Object Windows.Forms.WebBrowser
    $wb.ScrollBarsEnabled = $false; $wb.ScriptErrorsSuppressed = $true
    $wb.Width = $w; $wb.Height = $h
    $wb.Navigate([string]$p.file)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($wb.ReadyState -ne 'Complete' -and $sw.ElapsedMilliseconds -lt 20000) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 }
    $loaded = ($wb.ReadyState -eq 'Complete')
    Pump 150
    $zoomOk = $true
    try {
      $ax = $wb.ActiveXInstance
      [void]$ax.GetType().InvokeMember('ExecWB', [Reflection.BindingFlags]::InvokeMethod, $null, $ax, @(63, 2, [int](100 * $scale), $null))
    } catch { $zoomOk = $false }
    Pump $settle
    $js = @'
(function(){var d=document,e=d.documentElement,b=d.body,cs=function(x){try{return window.getComputedStyle(x).backgroundColor;}catch(z){return "";}};
var bg=b?cs(b):"";if(!bg||bg==="transparent"||/rgba\(.*,\s*0\)/.test(bg)){bg=cs(e);}
var sw=e.clientWidth,sh=e.clientHeight,all=b?b.getElementsByTagName("*"):[],i,r,p,clip;
for(i=0;i<all.length&&i<8000;i++){r=all[i].getBoundingClientRect();if(r.width<=0||r.height<=0||(r.right<=e.clientWidth+2&&r.bottom<=e.clientHeight+2)){continue;}
clip=false;p=all[i].parentNode;while(p&&p!==b&&p.nodeType===1){try{if(window.getComputedStyle(p).overflow!=="visible"){clip=true;break;}}catch(z){}p=p.parentNode;}
if(clip||window.getComputedStyle(all[i]).position==="fixed"){continue;}if(r.right>sw){sw=Math.round(r.right);}if(r.bottom>sh){sh=Math.round(r.bottom);}}
return [d.documentMode,b?(b.getAttribute("data-ih-ready")||""):"",e.clientWidth,e.clientHeight,sw,sh,bg,(e.getAttribute("data-ih-err")||"")].join("\u0001");})()
'@
    $info = $null
    try { $info = [string]$wb.Document.InvokeScript('eval', @($js)) } catch {}
    $r = [ordered]@{ n = $p.n; loaded = $loaded; zoomOk = $zoomOk; docMode = 0; ready = $false; clientW = 0; clientH = 0; scrollW = 0; scrollH = 0; bg = '#FFFFFF'; errors = ''; png = $p.png; evalOk = [bool]$info; perf = '' }
    if ($info) {
      $f = $info -split [char]1
      $r.docMode = [int]$f[0]; $r.ready = ($f[1] -eq '1')
      $r.clientW = [int]$f[2]; $r.clientH = [int]$f[3]; $r.scrollW = [int]$f[4]; $r.scrollH = [int]$f[5]
      if ($f[6] -match 'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)') { $r.bg = '#{0:X2}{1:X2}{2:X2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3] }
      $r.errors = $f[7]
    } else {
      # the page's own script may be broken so badly that eval is unavailable: read what the DOM offers
      try { $r.errors = [string]$wb.Document.GetElementsByTagName('html')[0].GetAttribute('data-ih-err') } catch {}
    }
    # large things that fade or move are what makes a page stutter in the slide show: list them
    $perfJs = @'
(function(){try{var e=document.documentElement,all=document.body.getElementsByTagName("*"),big=e.clientWidth*e.clientHeight*0.125,out=[],n=0,inf=0,i,el,cs,r,d,prop,an,name;
function dur(t){var m=0;String(t||"").split(",").forEach(function(x){var v=parseFloat(x)||0;if(/ms/.test(x)){v=v/1000;}if(v>m){m=v;}});return m;}
for(i=0;i<all.length&&i<6000;i++){el=all[i];cs=window.getComputedStyle(el);if(!cs){continue;}
prop=String(cs.transitionProperty||"");d=dur(cs.transitionDuration);an=!!cs.animationName&&cs.animationName!=="none"&&dur(cs.animationDuration)>0;
if(!an&&(d<=0||!/all|opacity|transform|width|height|left|top|right|bottom|margin|padding|box-shadow|filter/.test(prop))){continue;}
r=el.getBoundingClientRect();if(r.width*r.height<big){continue;}n++;if(an&&/infinite/.test(String(cs.animationIterationCount||""))){inf++;}
if(out.length<4){name=el.tagName.toLowerCase()+(el.id?"#"+el.id:"")+(typeof el.className==="string"&&el.className?"."+el.className.replace(/^\s+|\s+$/g,"").replace(/\s+/g,"."):"");
out.push(name+" "+Math.round(r.width)+"x"+Math.round(r.height)+(an?" animation":" transition "+prop));}}
return n+"|"+inf+"|"+out.join("; ");}catch(x){return "";}})()
'@
    try { $r.perf = [string]$wb.Document.InvokeScript('eval', @($perfJs)) } catch { $r.perf = '' }
    if ($A.probe) {
      $pv = ''
      try { $pv = [string]$wb.Document.InvokeScript('eval', @('(function(){try{var v=(' + [string]$A.probe + ');return typeof v==="string"?v:JSON.stringify(v);}catch(e){return "PROBE ERROR: "+e.message;}})()')) } catch { $pv = 'PROBE FAILED: ' + $_.Exception.Message }
      $r.probe = $pv
    }
    $bmp = New-Object Drawing.Bitmap($w, $h)
    $wb.DrawToBitmap($bmp, (New-Object Drawing.Rectangle(0, 0, $w, $h)))
    $bmp.Save([string]$p.png, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose(); $wb.Dispose()
    L ("page {0}: {1}x{2} canvas, mode={3} ready={4} view={5}x{6} scroll={7}x{8} bg={9} errors=[{10}]" -f $p.n, $p.canvasW, $p.canvasH, $r.docMode, $r.ready, $r.clientW, $r.clientH, $r.scrollW, $r.scrollH, $r.bg, $r.errors)
    $out += [pscustomobject]$r
  }
  Write-JsonFile $A.result ([ordered]@{ pages = $out })
} catch {
  L "ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"
  Write-JsonFile $A.result ([ordered]@{ error = $_.Exception.Message })
}
