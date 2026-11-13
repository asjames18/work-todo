# Work To-Do: read-only bridge to classic Outlook (COM). Called by todo.pyw with a timeout.
# It NEVER sends, moves, deletes, flags, marks read or saves anything. It only reads Inbox headers + a short preview.
# Inputs come from environment variables (never pasted into the script): WTD_ACTION, WTD_DAYS, WTD_SCAN, WTD_ENTRYID, WTD_MSGPATH.
# Output: one JSON line on stdout.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
function Out-Json($o) { [Console]::Out.WriteLine(($o | ConvertTo-Json -Depth 5 -Compress)) }
function Clip($s, $n) { if ($null -eq $s) { return '' }; $s = [string]$s; if ($s.Length -gt $n) { return $s.Substring(0, $n) }; return $s }

$action = $env:WTD_ACTION
if (-not $action) { $action = 'status' }

# 1) Checks that never start Outlook: is it installed / registered, and is there a mail profile?
$exe = $null
foreach ($p in @("$env:ProgramFiles\Microsoft Office\Root\Office16\OUTLOOK.EXE", "${env:ProgramFiles(x86)}\Microsoft Office\Root\Office16\OUTLOOK.EXE",
                 "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE", "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE")) {
  if ($p -and (Test-Path -LiteralPath $p)) { $exe = $p; break }
}
$registered = Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID'
$profiles = @()
$pk = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
if (Test-Path -LiteralPath $pk) { $profiles = @(Get-ChildItem -LiteralPath $pk -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName }) }
$running = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
$newOutlook = [bool](Get-Process -Name olk -ErrorAction SilentlyContinue)
$base = @{ exe = $exe; registered = $registered; profiles = $profiles.Count; running = $running; new_outlook_running = $newOutlook }

if (-not $exe -or -not $registered) { $base.state = 'not_installed'; Out-Json $base; exit 0 }
if ($profiles.Count -eq 0) { $base.state = 'no_profile'; Out-Json $base; exit 0 }
if ($action -eq 'status') { $base.state = 'ready'; Out-Json $base; exit 0 }

# 2) COM (this starts classic Outlook in the background if it isn't running; there is a profile at this point)
try {
  $ol = New-Object -ComObject Outlook.Application
  $ns = $ol.GetNamespace('MAPI')
} catch {
  $base.state = 'busy'; $base.error = Clip $_.Exception.Message 200; Out-Json $base; exit 0
}

try {
  if ($action -eq 'fetch') {
    $days = 7; if ($env:WTD_DAYS -match '^\d{1,2}$') { $days = [int]$env:WTD_DAYS }
    $scan = 300; if ($env:WTD_SCAN -match '^\d{1,4}$') { $scan = [int]$env:WTD_SCAN }
    $me = @{ name = ''; smtp = @(); x500 = '' }
    try { $me.name = [string]$ns.CurrentUser.Name } catch {}
    try { $me.x500 = [string]$ns.CurrentUser.Address } catch {}
    try { foreach ($a in $ns.Accounts) { if ($a.SmtpAddress) { $me.smtp += [string]$a.SmtpAddress } } } catch {}
    try { $xu = $ns.CurrentUser.AddressEntry.GetExchangeUser(); if ($xu -and $xu.PrimarySmtpAddress) { $me.smtp += [string]$xu.PrimarySmtpAddress } } catch {}
    $inbox = $ns.GetDefaultFolder(6)   # olFolderInbox, read only
    $since = (Get-Date).ToUniversalTime().AddDays(-$days).ToString('yyyy-MM-dd HH:mm')
    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $true)
    $found = $items.Restrict("@SQL=""urn:schemas:httpmail:datereceived"" >= '$since'")
    $found.Sort('[ReceivedTime]', $true)
    $out = @(); $n = 0
    foreach ($it in $found) {
      if ($n -ge $scan) { break }; $n++
      try {
        if ($it.Class -ne 43) { continue }   # olMail only
        $rc = @()
        $i = 0
        foreach ($r in $it.Recipients) { if ($i -ge 40) { break }; $i++; $rc += @{ t = [int]$r.Type; n = (Clip $r.Name 120); a = (Clip $r.Address 300) } }
        $from = ''
        try { if ($it.SenderEmailType -eq 'EX') { $xs = $it.Sender.GetExchangeUser(); if ($xs) { $from = [string]$xs.PrimarySmtpAddress } } } catch {}
        if (-not $from) { $from = [string]$it.SenderEmailAddress }
        $out += @{
          entry_id = [string]$it.EntryID; subject = (Clip $it.Subject 300); sender = (Clip $it.SenderName 160); sender_email = (Clip $from 300)
          received = $it.ReceivedTime.ToString('s'); to = (Clip $it.To 600); cc = (Clip $it.CC 600); importance = [int]$it.Importance
          unread = [bool]$it.UnRead; recips = $rc; body = (Clip $it.Body 1500)
        }
      } catch { }
    }
    Out-Json @{ state = 'ok'; me = $me; items = $out; scanned = $n; days = $days }
  } elseif ($action -eq 'open') {
    $id = $env:WTD_ENTRYID
    if ($id -notmatch '^[0-9A-Fa-f]{20,1024}$') { Out-Json @{ state = 'error'; error = 'bad id' }; exit 0 }
    $m = $ns.GetItemFromID($id)
    $m.Display()   # opens the email in its own Outlook window; changes nothing
    Out-Json @{ state = 'ok' }
  } elseif ($action -eq 'msg') {
    $path = $env:WTD_MSGPATH
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { Out-Json @{ state = 'error'; error = 'file not found' }; exit 0 }
    $m = $ns.OpenSharedItem($path)
    Out-Json @{ state = 'ok'; subject = (Clip $m.Subject 300); sender = (Clip $m.SenderName 160); sender_email = (Clip $m.SenderEmailAddress 300)
                received = $(try { $m.ReceivedTime.ToString('s') } catch { '' }); to = (Clip $m.To 600); cc = (Clip $m.CC 600); body = (Clip $m.Body 1500) }
    try { $m.Close(1) } catch {}   # olDiscard: never saves
  } else {
    Out-Json @{ state = 'error'; error = 'unknown action' }
  }
} catch {
  $msg = Clip $_.Exception.Message 200
  $st = 'error'
  if ($msg -match '80010001|8001010A|RPC_E_CALL_REJECTED|busy') { $st = 'busy' }
  Out-Json @{ state = $st; error = $msg }
}
