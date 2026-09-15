# Signature Add-in (Edelburg and LAF tenants)

Outlook add-in that auto-inserts a mailbox's signature when composing,
replying to, or forwarding a message — in classic Outlook for Windows, new
Outlook for Windows, Outlook on the web, and Outlook on Mac.

It exists because Exchange Online's own roaming-signature mechanism
(`Set-MailboxMessageConfiguration`) doesn't reliably populate the
`SignaturesList` that these clients actually read from — see the
`signatures` repo's history for that investigation. This add-in sidesteps
it entirely: it fetches the mailbox's pre-rendered signature HTML from
`https://s3-signatures.edelburg.net/<domain>/<local>.html` (built by that repo's
`Generate-Signatures.ps1`) and inserts it directly via the Office.js
`setSignatureAsync` API.

Mailboxes with more than one signature choose another from the side panel
(see "Signature panel" below). The same code serves two Microsoft 365
tenants, each with its own manifest (see "Tenants" below).

## How it identifies the sender

`functions.js` and `taskpane.js` read
`Office.context.mailbox.userProfile.emailAddress` (the signed-in mailbox's
own address — no external lookup) and turn it into a path of domain folder
and local part, lowercased: `J.Doe@example.com` →
`example.com/j.doe`. The `signatures` repo's generator writes its
output under the same paths, so the two must agree. The domain folder is
what lets one bucket serve both tenants: `j.doe@example.com` and
`j.doe@example.org` never collide.

Trade-off, discussed and accepted: these paths are guessable by anyone who
knows someone's address, unlike the opaque per-mailbox GUID approach
considered earlier. The bucket's Public Access is enabled but listing is
disabled, so it's not browsable/crawlable — just not resistant to someone
deliberately guessing a specific address.

## Files

- `manifest.xml` — the add-in manifest (classic XML, not the unified
  Microsoft 365 JSON manifest). References `functions.html`/`functions.js`
  and the icons below by URL. This is the Edelburg tenant's.
  `manifest.laf.xml` is the LAF tenant's, and `manifest.dev.xml` a copy of
  the Edelburg one with its own `Id` and name, for sideloading next to the
  deployed one. Keep all three in step (see "Tenants").
- `functions.html` / `functions.js` — the actual logic. `checkSignature()`
  runs automatically via a `LaunchEvent` on `OnNewMessageCompose` (fires for
  new messages, replies, *and* forwards) — no task pane or button involved.
  Every step is logged to the console for diagnosability, since failures
  are deliberately silent to the end user (see Known caveats).
- `taskpane.html` / `taskpane.js` — the side panel for choosing another
  signature (see "Signature panel").
- `assets/icon-{16,32,80,128}.png` — Edelburg's icons for the add-in listing
  and the ribbon button; `assets/laf/` holds LAF's. Edelburg's are generated to match the signature's palette (`#17241e` dark
  green, `#ab9569` gold, `#f4efe1` cream).

An earlier version had a manual "Insert Signature" ribbon button that was
dropped while `functions.html` wasn't loading at all (see git history).
Once the automatic path was confirmed working end to end, a ribbon button
came back: first as a static company menu, then as the Signatur panel, so
that new signatures don't need a manifest change.

## Signature panel

The compose ribbon gets a button (**Edelburg → Signatur**, or **LAF →
Signature**; in Outlook on the web and new Outlook it may sit under the
**Apps** button). It opens `taskpane.html`, a side panel that:

1. derives the mailbox's path the same way `functions.js` does and fetches
   `<domain>/<local>.json` from R2, which the `signatures` repo's generator
   writes for every mailbox, listing its signatures with a label and a
   file name and marking the standard one;
2. fetches each listed `.html` and shows it as a scaled-down preview
   (rendered in a shadow root, so the panel's CSS can't restyle it);
3. on click, inserts that HTML with `setSignatureAsync`, which replaces
   the signature already in the message.

Everyone sees only their own mailbox's signatures. Mailboxes without a
`.json` get a short "no signatures" note.

Because the list is data on R2 rather than part of the manifest, adding a
company or giving someone another signature is done entirely in the
`signatures` repo: regenerate, upload the new `.html`, the changed
`.json` files and any new logo. Nothing here changes and there is no
Microsoft 365 rollout delay. Only changes to the button itself (its
label, icon or the pane URL) need a manifest update.

The panel can't be tested outside Outlook as-is, since it needs
`office.js`. For a browser check, serve the page with a stub that defines
`Office.context.mailbox.userProfile` and `item.body.setSignatureAsync`,
and rewrites the R2 URLs to local `dist/` copies.

## Tenants

| Tenant | Manifest | Name | Panel language | Automatic signature |
|---|---|---|---|---|
| Edelburg | `manifest.xml` | Edelburg Signature | German | yes, on compose |
| LAF | `manifest.laf.xml` | LAF Signatures | French | no, panel only |

Both manifests load the same `functions.html`, `functions.js` and
`taskpane.html` from GitHub Pages and read the same bucket; the domain
folder in the path keeps their mailboxes apart. They differ only in `Id`,
display name, icons, ribbon strings, the task pane URL's `?lang=` (`de` is
the default when absent) and whether they insert a signature on compose:
`manifest.laf.xml` has no `LaunchEvent` (and no `Runtimes`), so LAF
mailboxes only ever get the signature picked in the panel, and its pane URL
adds `auto=0` so the panel doesn't badge one as automatic. Each is uploaded once in its own
tenant's Integrated Apps.

A push here goes live in both tenants at once, which is fine while one
person administers both. If that changes, publish released versions under
fixed paths and point each manifest at its own.

Adding a tenant: copy `manifest.laf.xml` with a new `Id`, name, icons and
language (add the language to `STRINGS` in `taskpane.js` if it's new),
upload it in that tenant, and add its mailboxes to the `signatures` repo.

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
in five places (`IconUrl`, `HighResolutionIconUrl`, the legacy
`SourceLocation`, and the two `bt:Url` entries under `Resources`). If your
repo ends up somewhere else, update all five before deploying — a
mismatched URL there means Outlook can't load the icon or the runtime and
the add-in silently fails to activate.

### 3. Test before rolling out

Sideload for a single-mailbox test: Outlook → **Get Add-ins → My add-ins →
Add a custom add-in → Add from file** → select your local `manifest.xml`.
Open a new message and confirm the signature appears automatically; try a
reply and a forward too.

If nothing appears: open the browser devtools console *before* composing,
clear it, then open the compose window. Check the frame/context selector at
the top of the Console panel — `functions.html` runs in its own frame
(`edelburg-kg.github.io`), and by default the console only shows the host
page's (e.g. `outlook.cloud.microsoft`) output. If that frame isn't listed
in the selector at all, `functions.html` never loaded — a manifest/plumbing
problem, not a bug inside `functions.js`. If it is listed and you switch to
it, every step logs (see `functions.js`), so a fetch failure, CORS block,
or `setSignatureAsync` error will show up explicitly rather than as silence.

### 4. Deploy to everyone

Microsoft 365 admin center → **Settings → Integrated apps → Upload custom
apps**. When asked how to provide the manifest, choose **"Provide link to
file"** and give it
`https://edelburg-kg.github.io/outlook-signature-addin/manifest.xml`,
rather than uploading the file directly. Both work today, but the URL
option means a future manifest change (version bump, a new button, a
permission tweak) just needs a `git push` — Outlook re-fetches the manifest
from that URL periodically, no re-upload through the admin center required.
A direct file upload is a frozen snapshot: it won't pick up edits to
`manifest.xml` until someone re-uploads it by hand.

Target **Entire organization** or specific users/groups, and deploy. Allow
up to 72 hours to propagate per Microsoft's own guidance.

## Known caveats

- **The signature HTML must not reference an SVG image.** Outlook on the
  web rejects *every* body write (`setSignatureAsync`, `prependAsync`,
  `setAsync`) whose HTML contains `<img src="….svg">`, and the only error
  it gives is the generic `Host Error` / code 5000 ("The operation is not
  supported") — nothing points at the image. The Office.Body API notes say
  "SVG files aren't supported in mail signatures. Use JPG or PNG files
  instead", and Microsoft confirmed this as the cause of the identical
  symptom in `OfficeDev/office-js#6020`. This add-in's first version shipped
  with the logo as an SVG and failed deterministically for days across
  sideloaded and admin-deployed installs before that was found. The
  `signatures` repo now renders the logo to PNG; `functions.js` logs a
  `console.error` if a fetched signature ever references an `.svg` again.
- **Fallback path.** If `setSignatureAsync` still fails after its retries,
  `functions.js` falls back to `body.prependAsync`, a plain body write
  supported since Mailbox 1.1. It lacks replace-on-reinsert semantics and
  dirties the form, so it should never be the path that runs in practice;
  the console's final `done:` line says which path did.

- **`OnNewMessageCompose` event-based activation has had reported reliability
  gaps on some Outlook desktop builds** (tracked upstream in
  `OfficeDev/office-js`) — occasionally it doesn't fire on a fresh compose
  window. The Signatur panel's standard signature doubles as a manual
  fallback.
- **`setSignatureAsync` can duplicate rather than replace** a signature if
  the same draft is opened across different Outlook platforms in sequence
  (tracked as `OfficeDev/office-js#5483`, unresolved upstream). Edge case,
  not expected to affect a fresh compose in one client.
- **No signature for mailboxes without a matching `<domain>/<local>.html`**
  (a shared mailbox, a guest, a new hire not yet in `signatures.csv`) — the fetch
  fails and `functions.js` leaves the compose window untouched rather than
  showing an error, since there's nothing actionable for the user to do
  about it in that moment.
