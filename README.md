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
- Needs both `Code.gs` (new `api_flags` and `api_action`) and `Index.html` updated in Apps Script.
- **Migrate quote:** a Confirmed work whose quote exists only as a Google Site still counts as quoted, but shows an amber warning, because outside vendors are picked automatically only from GitHub quotes. The **Migrate quote** chip lists them. The warning goes away once a GitHub quote is issued for that work.
