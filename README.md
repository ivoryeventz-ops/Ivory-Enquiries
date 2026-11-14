# Ivory Enquiries

Installable phone app for Ivory's enquiry book. `index.html` (repo root) is the thin wrapper that opens the Google Apps Script web app inside a frame; the icons and `manifest.webmanifest` make it installable.

## apps-script/

These two files are the Apps Script project attached to the **Ivory Enquiry Book** Google Sheet (Extensions > Apps Script). They are kept here as the source copy; deploying is done in Google.

- `Code.gs` - back end (enquiries, quotes, vendor ticks, Vendor list, reminders)
- `Index.html` - the app screen (create it in Apps Script as an HTML file named `Index`)

After changing either file in Apps Script: Deploy > Manage deployments > edit > New version.

## Vendor list > Not assigned

Opens from the **Vendor list** button. Lists, for every confirmed work, each outside-vendor item that has no vendor name yet (the 12 standard vendors plus anything under "Outside Vendors" in the latest quote), grouped by work and date. Decor package items are never listed, and items marked "Included in Package" are left out. Type a name next to an item and it saves to the Vendors tab straight away. No changes to `Code.gs` were needed for this view.

## No quote given + action reminders

- **No quote given:** a Confirmed work with no quote (neither a Google Site quote on the Quotes tab nor a page in ivory-packages) shows "No quote given yet" under its status. Use the **No quote given** chip to list them all.
- **Reminders:** each Confirmed work shows what is due and not yet ticked - 3 days before (wedding timeline, vendor checklist, payment & advance details), 1 day before (checklist to staff), day after (final payment details). Tick them on the card or inside the work; ticks are saved in a new **Actions** tab, shared by both phones. Use the **Action due** chip to see only works with something pending. A finished wedding stays on the list for 7 days while its day-after task is pending.
- **Signage board details:** a 3-days-before task, listed right after the wedding timeline. Ticks save the same way as the others.
- **Wedding timeline auto-tick:** the "Wedding timeline to be given" task shows ticked by default once a client page exists in the `ivory-wedding-timeline` repo. A folder counts when it shares a name word with the client and its page mentions the work's event date, so two weddings for one client never share a tick. A tick you save (or an un-tick) always overrides the default.
- Needs both `Code.gs` (new `api_flags` and `api_action`) and `Index.html` updated in Apps Script. The signage task and the timeline auto-tick are `Index.html` only.
- **Actions appear only when due:** inside a work, each stage (3 days before, 1 day before, day after) is shown only from the day it needs action - no "in 2 days" placeholders. Actions sit at the bottom of the work, below the quotes and outside vendors, so they never push the quotes down.
- **Opening a work is instant:** the last quotes/vendors shown for a work appear immediately and are refreshed quietly behind them (never redrawn while someone is typing or has an unsaved vendor change).
- **Vendors follow the latest quote:** opening a work or the Vendor list re-reads the quote pages (at most every 90 seconds), so a re-issued quote updates the outside-vendor list. "Anchor / MC", "LED Wall" and "Tall Men Security" are now standard outside vendors (Photo Booth is not), and "Ambient" lighting is recognised as well as "Ambiance". New vendor types need `Index.html` only; `Code.gs` also lists them.
- **Migrate quote:** a Confirmed work whose quote exists only as a Google Site still counts as quoted, but shows an amber warning, because outside vendors are picked automatically only from GitHub quotes. The **Migrate quote** chip lists them. The warning goes away once a GitHub quote is issued for that work.
- **Payment details:** inside a work, a **Payment details** row sits on top of *Quotes sent*, above the latest quote. It opens a sheet pre-filled from the latest quote in ivory-packages (name, date, venue, the decor package rate (from the package total on the page, or a package price line) and each additional service's rate). Rates, Total, Discount and Advance Received are all editable, and **+ Additional advance received** adds more advance lines (label such as "Advance Received on 23 Oct" plus amount). Balance is worked out live. **Save & publish** writes `<Client>/payment.html` in the packages repo (same folder as the client's quotes; page format matches `Christy/payment`), then offers **Send on WhatsApp** and **Open page**. Opening it again loads what was saved (older hand-made payment pages are read too). A *Sent* tag shows on the row once a payment page exists. The GitHub key (fine-grained, Contents read/write on Ivory-Packages) is asked for once, the first time anyone publishes, and kept in the Apps Script project's Script Properties (`GH_KEY`), so both phones use it and nobody is asked again; a rejected key is cleared and asked for again. Needs `Code.gs` (new `api_pay_status` and `api_pay_publish`) and `Index.html`; after pasting `Code.gs`, run any function once from the editor (or just publish) and approve the "connect to an external service" permission. With an old `Code.gs` the app falls back to a key stored on each phone.
