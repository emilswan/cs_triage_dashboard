# CS Triage Dashboard

Live: https://cs-triage-dashboard.vercel.app (login: user `cs`, shared password held in Vercel)

Read-only dashboard over the **customerservice@central-pharma.com** Inbox. Shows incoming mail by priority, who each message is assigned to (**assigned = internal CS people in the To line**), and whether it has been picked up (**picked up = a reply sent from the CS mailbox on the same conversation**, detected via Sent Items; an Outlook category still counts as a manual override). Critical/high items with no reply yet are listed for triage.

> Reply attribution: if reps "send as" the shared mailbox, Graph can't see who clicked send — the dashboard then credits the category or the assignee. "Send on behalf" preserves the individual sender name.

## How it works

```
Browser (public/index.html)
   │  Basic auth (middleware.js) — fails closed if DASHBOARD_PASSWORD unset
   ▼
/api/messages  ──client credentials──▶ Microsoft Graph ──▶ CS mailbox Inbox
/api/triage    ──(optional)──────────▶ Anthropic API (priority classification)
```

- `public/index.html`: the page (Chart.js from jsdelivr). Range presets, filters, auto-refresh every 5 min.
- `api/messages.js`: Graph `GET /users/{MAILBOX}/mailFolders/inbox/messages`, filtered by date, up to 20×100 messages per request.
- `api/triage.js`: only active if `ANTHROPIC_API_KEY` is set. Otherwise the page uses keyword rules. Sends sender, subject and a 300-character preview per email.
- `middleware.js`: Basic auth over the whole site.

## Hosting

- Vercel team `es-projects-bc33dd9c`, project `cs-triage-dashboard`, functions region **lhr1** (London).
- Not Git-connected: deployed directly from files. To update, redeploy the files (or connect a repo).

## Environment variables (Vercel → Settings → Environment Variables)

| Key | Notes |
|---|---|
| `MAILBOX` | customerservice@central-pharma.com |
| `MS_TENANT_ID` | 4c449397-8cec-4b2c-9ddf-f08152331829 |
| `MS_CLIENT_ID` | 3c024111-b3fc-4947-9ee6-54a809068231 (Entra app "CS Triage Dashboard") |
| `MS_CLIENT_SECRET` | Sensitive. Client secret of the app above. **Note its expiry date in Entra and rotate before it lapses.** |
| `DASHBOARD_PASSWORD` | Sensitive. Shared team password. `DASHBOARD_USER` optional (default `cs`). |
| `ANTHROPIC_API_KEY` | Optional. Enables AI triage. `ANTHROPIC_MODEL` optional (default `claude-haiku-4-5`). |

Changing a variable requires a **redeploy** to take effect.

## Microsoft permissions

The app has **no** Graph API permissions in Entra (deliberately: tenant-wide Mail.Read would expose every mailbox). Access is granted with Exchange RBAC for Applications, scoped to the CS mailbox only:

```powershell
New-ServicePrincipal -AppId 3c024111-b3fc-4947-9ee6-54a809068231 -ObjectId b4dbcb1b-ca16-4d80-adf5-4d8f0570a77f -DisplayName "CS Triage Dashboard"
New-ManagementScope -Name "CS mailbox only" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'customerservice@central-pharma.com'"
New-ManagementRoleAssignment -App 3c024111-b3fc-4947-9ee6-54a809068231 -Role "Application Mail.Read" -CustomResourceScope "CS mailbox only"
Test-ServicePrincipalAuthorization -Identity 3c024111-b3fc-4947-9ee6-54a809068231 -Resource customerservice@central-pharma.com   # InScope = True
```

Running these requires Exchange Administrator / Global Administrator.

## Known limits / next steps

- One shared password; Entra ID sign-in per user would be stronger.
- "Open ↗" links open Outlook on the web, so they only work for people with access to the CS mailbox.
- Customer names come from a sender-domain map (`CUSTOMER_NAMES` in index.html). Add new customers there.
- The AI triage cache is per browser (localStorage).
