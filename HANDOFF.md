# HANDOFF: Copy Email Link (Outlook add-in)

Written for an AI assistant (or developer) picking this project up cold. Read this first, then `manifest.xml`, `link.js`, `taskpane.js`.

## Goal
Let the owner, on **classic Outlook for Windows** (Microsoft 365 / Exchange Online mailbox), copy a link to the selected email that **opens that email in Outlook desktop**, so it can be pasted into other tools (tickets, notes, chats). A plain Outlook-on-the-web link was built first, but the owner wants links that open in the desktop app.

## Current state (as of the last session)
* Add-in is a static site on **GitHub Pages**: `https://pupontech.github.io/outlook-copy-link/` (repo `pupontech/outlook-copy-link`, branch `main`, root).
* Manifest v1.2.0.0, installed in classic Outlook via *Get Add-ins > My add-ins > Custom add-ins > Add from file*. ("Add from URL" was blocked or failed for the owner.)
* One ribbon button, **Link Panel** (Home tab, group "Email Link"), opens a pinnable task pane.
* The pane shows:
  * **Desktop link** `outlook:<HexEntryId>` (main box). Buttons: *Copy link*, *Copy as hyperlink*.
  * **Web link** (collapsed `<details>`): `https://outlook.office365.com/owa/?ItemID=<EWS id, URL-encoded>&exvsurl=1&viewmodel=ReadMessageItem`, button *Copy web link*.
* **Not yet verified in real Outlook:** that the `outlook:` link actually opens the message when clicked. v1.2 was only tested with a mocked Office.js. Verifying this is the first thing to do.
* The owner reported the **web** link opened in Outlook Online and "errored" when clicked (exact error not captured; likely signed in to a different account or tenant in the browser, or the item not found for that account). Not diagnosed.

## How the link is built
1. `Office.context.mailbox.item.itemId` (read mode) is an **EWS id** (standard base64, starts `AAMk...`). Web link = this id URL-encoded into the OWA deep link. This format matches the `webLink` Microsoft Graph returns.
2. Desktop link: `link.js > entryIdHex()` calls `Office.context.mailbox.makeEwsRequestAsync` with an EWS `ConvertId` (`SourceFormat` EwsId, `DestinationFormat="HexEntryId"`, `Mailbox` = `userProfile.emailAddress`), parses the `Id` from the returned `AlternateId`, and prefixes `outlook:`.
3. This needs manifest permission **ReadWriteMailbox** (Outlook shows a consent prompt on install).
4. `outlook:<EntryID>` is the classic Outlook URL protocol (registered as `OUTLOOK.EXE /select "%1"`). It does **not** work in New Outlook, Outlook on the web or Mac, and only on a PC with classic Outlook open to the same mailbox.

## Hard-won findings (do not re-learn these)
* **Clipboard:** classic Outlook blocks copying from ribbon *function commands* (headless runtime) and from a task pane that tries to copy automatically on load. Copying works only from a **click inside the task pane**. So a true one-click ribbon copy is impossible with an add-in; the `ExecuteFunction` "Copy Link" button and auto-copy were tried and removed. `commands.html`/`commands.js` are leftovers, not referenced by the manifest, and can be deleted.
* The owner's Outlook add-in task pane runs in **WebView2**. Computer-use/screenshot tools cannot see its contents (the `msedgewebview2.exe` window is masked), but the clipboard *can* be read to verify a copy.
* A manifest installed from a file **does not auto-update**. After any manifest change (including version, permissions, buttons) the owner must remove the add-in and add the manifest again. Changes to `*.js`/`*.html` need no reinstall, but GitHub Pages caches ~10 minutes, so the HTML references scripts as `link.js?v=N` / `taskpane.js?v=N`. **Bump N** whenever those scripts change.
* Office.js must be loaded from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`.
* Scripts are written in ES5-style with Promises; classic Outlook needs the WebView2 runtime.

## Files
| File | Purpose |
|---|---|
| `manifest.xml` | Add-in manifest (XML, `MailApp`, VersionOverrides 1.1, Mailbox 1.5, `MessageReadCommandSurface`). Host is `pupontech.github.io/outlook-copy-link`. |
| `taskpane.html` / `taskpane.js` | The Link Panel UI and logic. Re-renders on `ItemChanged` (when pinned). |
| `link.js` | `buildLink`, `currentLink`, `entryIdHex`, `desktopLink`, `copyText`, `copyHyperlink`. Also exports for Node (`module.exports`). `OWA_BASE` constant selects the web host (use `https://outlook.live.com/owa/` for Outlook.com accounts). |
| `commands.html` / `commands.js` | Unused leftovers from the removed one-click button. |
| `assets/icon-*.png` | Ribbon/store icons (16/32/64/80/128). |
| `README.md` | **Out of date**: still describes the removed Copy Link button and web-only link. Update it. |

## How to deploy changes
* There is no build step. Edit files, commit to `main`, GitHub Pages redeploys in about a minute.
* The previous session had no working git token, so changes were pushed via the GitHub web UI (*Add file > Upload files*, replacing same-named files). Use `git push` if credentials are available.
* Edit `manifest.xml` only deliberately; bump `<Version>`; tell the owner to remove and re-add the add-in.
* Do not commit mailbox data, message IDs, email addresses or tokens. The repo is **public**.

## How to test
* **Offline:** run `taskpane.html` in headless Chromium (Playwright) with a stub `Office` object (stub `onReady`, `context.mailbox.item.itemId/subject`, `userProfile.emailAddress`, `makeEwsRequestAsync` returning a `ConvertIdResponse` with `AlternateId Format="HexEntryId" Id="..."`, `addHandlerAsync`), routing `**/office.js` to the stub. Assert the textarea values, clipboard contents and the failure path. This was done and passes.
* **Real Outlook:** open the panel, click *Copy link*, read the clipboard (it should be `outlook:` followed by a long hex string), paste it into a Word doc or the Windows Run box, click it, confirm the message opens.

## Suggested next steps
1. Verify the `outlook:` link in real Outlook. If Outlook says it cannot open the item, check that `Mailbox` in the ConvertId request matches the primary mailbox, that the item is in the primary mailbox (not a shared mailbox, archive or other account), and whether the EntryID from `ConvertId` equals the one the Outlook object model reports. Shared mailboxes may need `Mailbox` set to the shared address.
2. If `outlook:` links cannot be made reliable: fall back to a **classic Outlook VBA macro** on the Quick Access Toolbar (true one-click). It would read `Selection(1).EntryID` and copy `outlook:<EntryID>` using a `MSForms.DataObject`. Needs macros enabled; classic Outlook on one PC only.
3. Diagnose the web link error (ask for the exact message; check which account the default browser is signed in to).
4. Update `README.md`, delete `commands.*`, and consider adding a LICENSE.

## Working with the owner
* The owner uses classic Outlook for Windows on a Microsoft 365 mailbox and wants to paste these links into other tools.
* Ask before publishing anything new, adding public content or changing permissions. Keep explanations short and plain.
