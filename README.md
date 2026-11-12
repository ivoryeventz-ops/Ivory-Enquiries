# Ivory Enquiries

Installable phone app for Ivory's enquiry book. `index.html` (repo root) is the thin wrapper that opens the Google Apps Script web app inside a frame; the icons and `manifest.webmanifest` make it installable.

## apps-script/

These two files are the Apps Script project attached to the **Ivory Enquiry Book** Google Sheet (Extensions > Apps Script). They are kept here as the source copy; deploying is done in Google.

- `Code.gs` - back end (enquiries, quotes, vendor ticks, Vendor list, reminders)
- `Index.html` - the app screen (create it in Apps Script as an HTML file named `Index`)

After changing either file in Apps Script: Deploy > Manage deployments > edit > New version.

## Vendor list > Not assigned

Opens from the **Vendor list** button. Lists, for every confirmed work, each outside-vendor item that has no vendor name yet (the 12 standard vendors plus anything under "Outside Vendors" in the latest quote), grouped by work and date. Decor package items are never listed, and items marked "Included in Package" are left out. Type a name next to an item and it saves to the Vendors tab straight away. No changes to `Code.gs` were needed for this view.
