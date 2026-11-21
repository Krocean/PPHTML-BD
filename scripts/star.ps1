param(
  [switch]$Yes,                        # the user said yes in chat: star the repository (gh), or open its page
  [switch]$No,                         # the user said no (or did not want to answer): remember it, never ask again
  [switch]$NoOpen,                     # with -Yes and no usable gh: only print the address, do not open the browser
  [string]$Repo = 'Krocean/PPHTML-BD'  # owner/name on github.com
)
# Asks for a GitHub star once per PC, and only with the user's yes.
#   (no switch)  Check only, does nothing on GitHub. Prints one line:
#                  STAR: skip (...)   -> do not ask, just do the work
#                  STAR: ask (gh)     -> ask the user; on yes, -Yes stars the repo with the user's gh login
#                  STAR: ask (link)   -> ask the user; on yes, -Yes opens the repo page for the user to press Star
#   -Yes         Run only after the user clearly said yes. Prints STAR: done (...) or STAR: link <address>.
#   -No          Prints STAR: noted. The question is not asked again on this PC.
# The answer is kept in %LOCALAPPDATA%\ppt-html\star.json (date and answer only). Delete that file to be asked again.
# Starring is never a condition for using the skill. Plain ASCII on purpose (see lib\hd_core.ps1).
# Exit codes: 0 done (whatever the answer); 1 unexpected error.
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false) } catch {}
function Say([string]$Text) { [Console]::Out.WriteLine($Text) }

$exit = 0
try {
  if ($Repo -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') { throw "bad -Repo value: $Repo (expected owner/name)" }
  $url = "https://github.com/$Repo"
  $root = Join-Path $env:LOCALAPPDATA 'ppt-html'
  $mark = Join-Path $root 'star.json'

  function Save-Mark([string]$Answer, [string]$How) {
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Force $root | Out-Null }
    $o = [ordered]@{ repo = $Repo; answer = $Answer; how = $How; date = (Get-Date -Format 'yyyy-MM-dd') }
    [IO.File]::WriteAllText($mark, ($o | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
  }
  function Read-Mark {
    if (-not (Test-Path -LiteralPath $mark)) { return $null }
    try { return ([IO.File]::ReadAllText($mark) | ConvertFrom-Json) } catch { return $null }
  }
  # gh is usable when it is installed and logged in to github.com. Returns the login name or ''.
  function Get-GhLogin {
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if (-not $gh) { return '' }
    $login = & gh api user --jq .login 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $login) { return '' }
    return ([string]$login).Trim()
  }
  function Test-Starred {
    & gh api "user/starred/$Repo" --silent 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
  }

  if ($Yes -and $No) { throw 'use -Yes or -No, not both' }

  if ($No) {
    Save-Mark 'no' ''
    Say 'STAR: noted (the user said no; this PC will not be asked again)'
  } elseif ($Yes) {
    $login = Get-GhLogin
    $done = $false
    if ($login) {
      if (Test-Starred) {
        $done = $true
        Say "STAR: done (already starred $Repo as $login)"
      } else {
        & gh api -X PUT "user/starred/$Repo" --silent 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0 -and (Test-Starred)) {
          $done = $true
          Say "STAR: done (starred $Repo as $login)"
        } else {
          Say "note: gh could not star $Repo (token without permission?); falling back to the page"
        }
      }
    }
    if ($done) {
      Save-Mark 'yes' 'gh'
    } else {
      $opened = $false
      if (-not $NoOpen) { try { Start-Process $url; $opened = $true } catch {} }
      Save-Mark 'yes' 'link'
      if ($opened) { Say "STAR: link $url (opened in the browser; the user presses Star at the top right of the page)" }
      else { Say "STAR: link $url (give this address to the user; Star is at the top right of the page)" }
    }
  } else {
    $m = Read-Mark
    if ($m -and $m.answer) {
      Say ("STAR: skip (already answered '{0}' on {1})" -f $m.answer, $m.date)
    } else {
      $login = Get-GhLogin
      if ($login -and (Test-Starred)) {
        Save-Mark 'yes' 'already'
        Say "STAR: skip (already starred by $login)"
      } elseif ($login) {
        Say "STAR: ask (gh) - on yes, -Yes stars $Repo as $login"
      } else {
        Say "STAR: ask (link) - gh is missing or not logged in; on yes, -Yes opens $url"
      }
    }
  }
} catch {
  $exit = 1
  Say "FAILED: $($_.Exception.Message)"
}
exit $exit
