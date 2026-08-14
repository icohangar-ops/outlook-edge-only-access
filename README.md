# Reading Outlook records when conditional access allows Edge only

A field note on giving an AI coding agent (Claude Code) read access to an
Exchange Online mailbox inside a tenant that permits browser sign-in from
Microsoft Edge only — and blocks user consent to third-party OAuth apps.

Written up because the two controls involved look like one problem and are
not, which cost a lot of time to work out. The PowerShell that came out of it
is included.

## Quick start

Requires Windows, classic Outlook (Office 16) running and signed in, and
PowerShell 5.1+. Everything is read-only.

```powershell
# 1. What am I actually working with? Stores, folder sync state, cache window.
.\scripts\Get-MailboxDiagnostics.ps1

# 2. Sent Items stuck at zero? Promote it in the sync queue.
.\scripts\Sync-OutlookFolder.ps1 -FolderId 5

# 3. Aggregate everyone you have emailed, with first/last contact and volume.
.\scripts\Export-SentRecipients.ps1 -Since '2025-02-01' -OutCsv .\recipients.csv

# 4. Optional: bucket them using your own domain map.
.\scripts\Group-RecipientsByCategory.ps1 -InCsv .\recipients.csv `
    -MapPath .\map.psd1 -OutDir .\out -InternalDomain 'company.com'

# Find a file when you only remember roughly what it was about.
.\scripts\Find-MailAttachment.ps1 -SubjectPattern 'tolling' -SaveTo .\out
```

## Repository layout

```
src/OutlookMapiTools/     PowerShell module - MAPI connection, folder walking,
                          SMTP resolution, sync priority, cache-window checks
scripts/                  Task scripts built on the module
examples/                 Sample category map (illustrative domains only)
tests/                    Pester tests; Outlook-dependent ones self-skip
```

Tests need **Pester 5+** (`Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser`).
Windows ships Pester 3.4, which cannot run them.

```powershell
Invoke-Pester -Path .\tests
Invoke-Pester -Path .\tests -ExcludeTagFilter RequiresOutlook   # no Outlook needed
```

## The problem

Two **independent** controls sat in front of the mailbox:

1. **Conditional access — Edge-only sign-in.** The agent's browser automation
   drives Chrome and its own embedded browser. Microsoft accepts neither.
   Every attempt hung indefinitely at *"Trying to sign you in."*
2. **Third-party OAuth consent blocked at tenant level.** Even from a
   permitted browser, authorising a non-Microsoft application returned
   *"Need admin approval."*

Solving either one alone changes nothing. Admin consent does **not** lift
conditional access, and using an approved browser does **not** grant consent.
Diagnosing them as a single failure is the trap.

## What worked

### 1. Graph-based connector — the sanctioned cloud path

An RPA platform's integration service (UiPath Integration Service, in this
case) connects to Outlook through the **Microsoft Graph API** rather than by
driving a browser, so the Edge-only restriction stops applying to the data
path. It still needs a one-time interactive authorisation, which is done in
Edge, plus tenant admin consent for the connector application.

Scope it down before authorising. The connector defaults to a broad set;
this deployment kept only:

```
openid  offline_access  User.Read  Mail.Read  email  profile
```

with `Mail.Send`, `Mail.ReadWrite`, `Calendars.Read`, `Calendars.Read.Shared`
and `Mail.Read.Shared` removed. The resulting connection cannot send mail as
the user.

Two findings worth recording:

- **The OAuth flow must start and finish in the same browser.** Beginning the
  flow elsewhere and completing it in Edge fails with `Invalid path provided`
  — the callback loses its folder/session context.
- **Admin consent and conditional access are separate approvals.** You need
  both, and the interactive sign-in still has to happen in Edge.

### 2. Desktop MAPI — what actually produced the extract

Classic Outlook on the workstation was configured against the same mailbox.
**Desktop applications satisfy conditional access where browsers do not**, so
signing the profile in works normally. Outlook then caches the mailbox
locally and the records are readable directly through MAPI/COM — no cloud
API, no browser, and no data leaving the machine.

This is the route the extract came from: 4,700+ sent messages scanned and
1,000+ distinct recipients aggregated. It was chosen over the connector
because reading data *back out* of an RPA workflow means building and running
a visual automation, whereas local MAPI gives direct, scriptable access with
straightforward iteration over a whole folder.

See [`scripts/Export-SentRecipients.ps1`](scripts/Export-SentRecipients.ps1).

## Gotchas that cost real time

**The offline cache is 12 months by default.** Asking for 18 months of
history silently returns 12. Widening it is not scriptable in a locked-down
environment:

- `HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Cached Mode\SyncWindowSetting`
  applies to **newly created accounts only** — it will not retrofit an
  existing profile.
- `HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode\...`
  is the branch that *does* apply to existing accounts, and is typically
  write-protected.

The reliable route is the UI: *File → Account Settings → Account Settings →
double-click the account → "Mail to keep offline" → 24 months → restart.*

**Outlook syncs the Inbox to completion before other folders.** Sent Items
sat at 1 item for 25+ minutes while the Inbox climbed past 3,800. Polling
alone will not fix it. Making the folder *active* promotes it in the sync
queue:

```powershell
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$sent = $ns.GetDefaultFolder(5)              # olFolderSentMail

$ol.ActiveExplorer().CurrentFolder = $sent   # equivalent to clicking it
for ($i = 1; $i -le $ns.SyncObjects.Count; $i++) { $ns.SyncObjects.Item($i).Start() }
```

Sent Items went from 1 to 4,747 items immediately after.

**Delegate mailboxes are attached to the profile.** `GetDefaultFolder()`
reads the default store only, which is what you usually want — but confirm,
because other people's mail can be one `Stores` index away.

**Exchange recipients expose a legacy DN, not an SMTP address.** Resolve via
`AddressEntry.GetExchangeUser().PrimarySmtpAddress`, falling back to
`PropertyAccessor` on `PR_SMTP_ADDRESS`
(`http://schemas.microsoft.com/mapi/proptag/0x39FE001E`).

## Choosing between the two paths

| | Graph connector | Desktop MAPI |
|---|---|---|
| Admin consent needed | Yes | No |
| Runs unattended / scheduled | Yes | No — needs the workstation |
| Data leaves the machine | Yes | No |
| Good for bulk extraction | Awkward — output via workflow | Direct and scriptable |
| Coverage | Whole mailbox, server-side | Only what is cached offline |

Rule of thumb: **connector for recurring automation, MAPI for one-off bulk
extraction.**

## Worth requesting from IT

Admin consent for **Microsoft Graph Command Line Tools**
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`, delegated `Mail.Read`). It is a
Microsoft first-party application and its **device-code sign-in is completed
in a browser of the user's choosing** — so an Edge-only policy is satisfied
while the token is then usable from any script. That combination makes
mailbox extracts directly scriptable without a local Outlook cache or an RPA
workflow.

## Scope and caveats

- Read-only throughout. Nothing here sends, modifies or deletes mail.
- Tenant names, mailbox identities and connection identifiers have been
  generalised.
- Behaviour described was observed on Outlook (Office 16) against Exchange
  Online in August 2026. Microsoft changes these surfaces; verify before
  relying on specifics.
