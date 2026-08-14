# Case study: unlocking a locked-down mailbox for automation

How an RPA platform (UiPath Automation Cloud) turned "no programmatic access
to this mailbox, at all" into an approved, least-privilege capability — and
what it did and did not solve.

## The business problem

A company preparing a **government loan application and a capital raise** needed
a defensible answer to a simple question:

> Who have we approached for capital in the last 18 months, and when?

The answer existed only as ~4,700 sent emails across an 18-month window. The
manual alternative was someone reading a year and a half of mail and
maintaining a spreadsheet by hand — days of work, stale the moment it was
finished, and impossible to reproduce or audit when a lender asked how the
list was compiled.

This is a common shape of problem. The record exists; it is just trapped in a
mailbox in a form no one can query.

## Why it was blocked

The mailbox sat behind two **independent** controls:

1. **Conditional access permitted browser sign-in from Microsoft Edge only.**
   Automation tooling that drives any other browser cannot authenticate. Not
   "degraded" — it hangs at the sign-in screen indefinitely.
2. **The tenant blocked user consent to third-party OAuth applications.** Even
   from an approved browser, authorising any connector returned
   *"Need admin approval."*

Both controls are correct security posture. Neither was going to be relaxed
for a reporting task. And critically, **solving one does nothing for the
other** — which is what made this look unsolvable for a while.

The organisational reality mattered as much as the technical one: there was no
established, approved channel for programmatic mailbox access. Every ad-hoc
request would have been a fresh negotiation with IT.

## What UiPath actually solved

**It provided the first sanctioned programmatic channel to the mailbox.**

The key insight is architectural. UiPath Integration Service does not automate
a browser to reach Outlook — it calls the **Microsoft Graph API**. Once the
connection is authorised, the Edge-only restriction is simply not on the data
path any more. The browser constraint stops being a constraint.

That reframing is what made the request approvable. Instead of asking IT to
weaken a conditional-access policy, the ask became: *grant admin consent to a
named enterprise application, scoped to read mail only.* That is a routine,
auditable decision an administrator can reason about — and it was **approved
within a day**.

Three things made it approvable:

- **Least privilege, demonstrably.** The connector requests a broad default
  scope set. It was trimmed before authorising to
  `openid offline_access User.Read Mail.Read email profile` — with
  `Mail.Send`, `Mail.ReadWrite`, `Calendars.Read`, `Calendars.Read.Shared` and
  `Mail.Read.Shared` all removed. The resulting connection **cannot send mail
  as the user**, cannot modify anything, and cannot reach colleagues' mailboxes.
- **A named application in the tenant.** Consent attaches to an identifiable
  enterprise app that appears in the admin console, can be reviewed, and can be
  revoked centrally at any time. That is a very different proposition from a
  personal script holding a token.
- **An auditable trail.** Connection ownership, identity and scope are all
  visible in the platform rather than living in someone's PowerShell profile.

**The durable business value is the channel, not any single report.** The
approval is a standing capability: scheduled and unattended automation against
the mailbox is now possible without a fresh IT conversation each time —
recurring investor updates, board-pack inputs, periodic extracts. The
precedent, and the pattern for scoping the request, is reusable.

## What it did not solve, and why that matters

**The one-off bulk extract was ultimately done a different way** — through
Outlook desktop over MAPI/COM on the workstation, which is what the scripts in
this repository do.

The reason is worth being honest about, because it informs tool selection:

Reading data *back out* of an RPA workflow means building a visual automation,
running it on a robot, and routing its output somewhere retrievable. For a
one-shot aggregation across thousands of messages — with pagination,
deduplication, address resolution and classification — a script with direct
folder access is simply the better instrument. Local MAPI also keeps the data
on the workstation entirely.

**This is not a criticism of the platform. It is a statement about fit.** RPA
connectors are built for repeatable, scheduled, governed processes. Ad-hoc bulk
analysis is a different job.

## Outcome

| | |
|---|---|
| Messages processed | ~4,700 |
| Distinct recipients extracted | ~1,000 |
| Capital-provider contacts identified | ~330 across ~100 firms |
| Manual effort replaced | Days of reading, non-reproducible |
| Time to re-run | Minutes |
| Standing capability gained | IT-approved, read-only mailbox automation |

The deliverable was a categorised contact list with first-contact date, last
contact date and message volume per counterparty — reproducible on demand, and
explainable to a lender asking how it was compiled.

## Choosing between the two

| | Graph connector (UiPath) | Desktop MAPI |
|---|---|---|
| Requires admin consent | Yes — once | No |
| Scheduled / unattended | **Yes** | No — needs the workstation |
| Governed and centrally revocable | **Yes** | No |
| Data leaves the machine | Yes | **No** |
| Coverage | **Whole mailbox, server-side** | Only what is cached offline |
| Ad-hoc bulk extraction | Awkward | **Direct** |

**Rule of thumb:** the governed connector for anything recurring; local MAPI
for one-off bulk work on a single workstation.

## Transferable lessons

1. **Change the shape of the request, not the policy.** "Relax conditional
   access" fails. "Approve this named app, read-only" succeeds. An API-based
   integration lets you ask the second question instead of the first.
2. **Scope down before you ask.** Arriving with permissions already trimmed to
   the minimum, and able to say plainly that the connection cannot send mail,
   converts a security conversation into an approval.
3. **Diagnose stacked controls separately.** Two independent blocks presenting
   as one symptom cost the most time here. Identify each before attacking either.
4. **The approval is worth more than the report.** A single extract answers one
   question; a sanctioned channel answers every future one.
