# Edelburg Signature Add-in

Outlook add-in that auto-inserts an employee's signature when composing,
replying to, or forwarding a message — in classic Outlook for Windows, new
Outlook for Windows, Outlook on the web, and Outlook on Mac.

It exists because Exchange Online's own roaming-signature mechanism
(`Set-MailboxMessageConfiguration`) doesn't reliably populate the
`SignaturesList` that these clients actually read from — see the
`signatures` repo's history for that investigation. This add-in sidesteps
it entirely: it fetches the employee's pre-rendered signature HTML from
`https://s3-signatures.edelburg.net/<alias>.html` (built by that repo's
`Generate-Signatures.ps1`) and inserts it directly via the Office.js
`setSignatureAsync` API.

## How it identifies the sender

`functions.js` reads `Office.context.mailbox.userProfile.emailAddress` (the
signed-in user's own address — no external lookup) and derives the alias as
the local-part, lowercased (`e.debeckerremy@edelburg.com` → `e.debeckerremy`).
This matches the `Alias` column in `employees.csv` for every current row.
Trade-off, discussed and accepted: alias filenames are guessable by anyone
who knows an employee's name pattern, unlike the opaque per-mailbox GUID
approach considered earlier. The bucket's Public Access is enabled but
listing is disabled, so it's not browsable/crawlable — just not resistant to
someone deliberately guessing a specific person's filename.

## Files

- `manifest.xml` — the add-in manifest (classic XML, not the unified
  Microsoft 365 JSON manifest). References the four items below by URL.
- `functions.html` / `functions.js` — the actual logic. One shared script:
  `checkSignature()` runs automatically via a `LaunchEvent` on
  `OnNewMessageCompose` (fires for new messages, replies, *and* forwards);
  `insertSignatureCommand()` backs a manual "Insert Signature" ribbon button
  for when auto-insert didn't fire or the user cleared the signature.
- `assets/icon-{16,32,80,128}.png` — ribbon/store icons, generated to match
  the signature's palette (`#17241e` dark green, `#ab9569` gold, `#f4efe1`
  cream).

## Setup

### 1. Create the repo and enable Pages

This code has zero PII in it (unlike the `signatures` repo, which stays
private) — that's why it's fine to host publicly on GitHub Pages while the
per-employee HTML/logo moved to R2.

```bash
gh repo create Edelburg-KG/outlook-signature-addin --public --source=. --push
```

(No `gh` CLI available in the environment this was built in — if you don't
have it either: create the repo at github.com/organizations/Edelburg-KG/repositories/new
as **public**, then `git init && git add -A && git commit -m "Initial add-in" && git remote add origin <url> && git push -u origin main`.)

Then: repo **Settings → Pages → Source: Deploy from a branch → main / (root)**.
Give it a few minutes, then confirm `https://edelburg-kg.github.io/outlook-signature-addin/functions.html`
loads (empty page is correct — there's no visible UI).

### 2. Fix the URLs if you used a different repo name/org

`manifest.xml` assumes `https://edelburg-kg.github.io/outlook-signature-addin/`
in six places (`IconUrl`, `HighResolutionIconUrl`, the legacy
`SourceLocation`, and the three `bt:Url`/`bt:Image` entries under
`Resources`). If your repo ends up somewhere else, update all six before
deploying — a mismatched URL there means Outlook can't load the icons or the
runtime and the add-in silently fails to activate.

### 3. Test before rolling out

Sideload for a single-mailbox test: Outlook → **Get Add-ins → My add-ins →
Add a custom add-in → Add from file** → select your local `manifest.xml`.
Open a new message and confirm the signature appears; try the "Insert
Signature" button too; try a reply and a forward.

### 4. Deploy to everyone

Microsoft 365 admin center → **Settings → Integrated apps → Upload custom
apps**. You can upload `manifest.xml` directly as a file (no need to host
the manifest itself — only the URLs it points to need to be public), target
**Entire organization** or specific users/groups, and deploy. Allow up to 72
hours to propagate per Microsoft's own guidance.

## Known caveats

- **`OnNewMessageCompose` event-based activation has had reported reliability
  gaps on some Outlook desktop builds** (tracked upstream in
  `OfficeDev/office-js`) — occasionally it doesn't fire on a fresh compose
  window. The manual "Insert Signature" button exists specifically as the
  fallback for this.
- **`setSignatureAsync` can duplicate rather than replace** a signature if
  the same draft is opened across different Outlook platforms in sequence
  (tracked as `OfficeDev/office-js#5483`, unresolved upstream). Edge case,
  not expected to affect a fresh compose in one client.
- **No signature for mailboxes without a matching `dist/<alias>.html`** (a
  shared mailbox, a guest, a new hire not yet in `employees.csv`) — the fetch
  fails and `functions.js` leaves the compose window untouched rather than
  showing an error, since there's nothing actionable for the user to do
  about it in that moment.
