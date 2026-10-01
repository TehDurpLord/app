# Parts Inventory

A Google Sheet for your parts that emails you the link to reorder something when it runs low.
No app and no website: everything is one spreadsheet, plus a little code inside it that sends the emails.

## What's in the spreadsheet

- **Inventory**: one row per part. Only *Part Name*, *Quantity*, *Min Qty* and *Order Link* matter;
  *Part Number*, *Location* and *Notes* are there if you want them.
- When a part's **Quantity** drops to its **Min Qty** or below, the row turns yellow (red at 0) and an
  email goes out with the part's order link. You get one email per shortage. You only get another one
  after the part has been restocked and runs low again.
- **Ordered**: tick it once you've ordered more. Ticked parts drop off the Reorder tab and the reminder
  email, and the box unticks itself once the part has run low and been restocked.
- **Reorder**: the list of parts that need ordering, with their links. It updates by itself; don't type in it.
- **Settings**: who gets the emails, the reminder schedule, and a **Status** line that says whether the
  emails are on.
- **Alert Sent** (grey column): filled in when a part's email goes out. Clear it to get emailed about
  that part again.
- A **reminder email** goes out on weekday mornings at 8 with everything still low and not ticked as Ordered.

## Set it up

### 1. Get the spreadsheet

Start a blank Google Sheet ([sheets.new](https://sheets.new)); the setup in step 2 builds the tabs.
To have the layout first, upload [`template/parts-inventory.xlsx`](template/parts-inventory.xlsx) to
Google Drive and open it in Google Sheets instead.

### 2. Turn on the emails (once, about two minutes)

1. In the spreadsheet, choose **Extensions → Apps Script**.
2. Delete what's in the editor and paste in all of [`src/Code.js`](src/Code.js). On GitHub, the
   *Copy raw file* button next to *Raw* copies all of it.
3. Click **Save** (the disk icon). Next to **Run**, it should say `setup`.
4. Click **Run**. Google asks for permission the first time: **Review permissions**, pick your account,
   then on *Google hasn't verified this app* click **Advanced → Go to … (unsafe)** and **Allow**. That
   warning appears for any script you write yourself; this one only runs in your own account.
5. The log at the bottom says *Done* and a test email arrives. Close the Apps Script tab.

The script can edit this spreadsheet (and no others), send email from your account, run on its own
(it checks the sheet once an hour), and see your email address (to send the emails to you).

### 3. Add your parts

Type them into the **Inventory** tab, or paste them in from another list. Give each part a **Min Qty**
to get emails about it, and paste its product page into **Order Link**.

## Day to day

- **Used some:** lower the Quantity. **Restocked:** raise it. The Google Sheets phone app works fine for this.
- **Got an email:** click *Order from …*, place the order, then tick **Ordered**.
- **What do we need?** Look at the **Reorder** tab, or choose **Inventory → Email the reorder list now**.
  The **Inventory** menu appears at the top of the spreadsheet after you reload it.

## Sharing with coworkers

Click **Share** in the spreadsheet and add them as **Editors**. They update counts right in the sheet,
and the emails still go out from your account; they don't need to do anything with the code. To email
them too, add their addresses next to **Send emails to** in the Settings tab.

## Settings

| Setting | What it does |
| --- | --- |
| Status | Filled in by the script: *On*, *Off* (not set up yet), or *Problem* with the reason. |
| Send emails to | Who gets the emails. Separate several addresses with commas. |
| Email right away when a part runs low | *On*: one email the moment a part runs low. *Off*: only the reminder. |
| Reminder email | *Off*, *Daily*, *Weekdays* or *Weekly* (Mondays). |
| Reminder hour (0-23) | When the reminder goes out, in the spreadsheet's time zone (File → Settings). 8 is 8 AM. |

## Use a spreadsheet you already have

Rename the tab with your parts to **Inventory** and do step 2. Your column headings are recognized in
common forms:

| Column | Also recognized as |
| --- | --- |
| Part Name | Name, Item, Part, Description |
| Part Number | Part #, Part No, PN, SKU, MPN, Item #, Model |
| Quantity | Qty, Qty On Hand, On Hand, In Stock, Stock, Count |
| Min Qty | Min, Minimum, Reorder Point, Reorder Level, Par, Threshold, Safety Stock |
| Order Link | Link, URL, Product Link, Reorder Link, Purchase Link, Website |
| Location | Bin, Shelf, Storage, Area |
| Ordered | Ordered? (tick boxes, or *Yes*) |
| Supplier, Reorder Qty, Unit | Vendor · Order Qty · UOM (shown in the emails when you have them) |

Setup adds the columns it needs at the end (for example *Ordered* and *Alert Sent*) and never deletes
or moves anything. A title row above the headings is fine, a *Total* row under the list isn't treated
as a part, and links made with *Insert → Link* or `=HYPERLINK()` work as order links. *Alert Sent* is
only recognized by that exact name, so a column of yours like *Alerted* is never written to.

## Troubleshooting

- **No emails:** check spam first. Then look at **Status** in the Settings tab and the **Send emails to**
  row, and try **Inventory → Send a test email**.
- **Status says Off:** the code isn't set up yet. Do step 2.
- **A part is low but no email came:** it was probably already emailed about in this shortage (its *Alert
  Sent* is filled in). Clear *Alert Sent* to send it again. A part with no *Min Qty* never sends emails.
- **No Inventory menu:** reload the spreadsheet. The emails work without the menu.
- **Don't rename** the *Inventory* or *Settings* tabs; the script finds them by name.
- **Someone else set it up and has left:** run **Inventory → Set up / repair** (or step 2) yourself. If
  two people have run setup, each email still goes out once.

---

## For developers

```
src/Code.js                    The whole script: setup, triggers, emails (paste this into Apps Script)
src/appsscript.json            Apps Script manifest (for clasp)
template/parts-inventory.xlsx  The ready-made spreadsheet, built from Code.js by npm run build
dev/build-sheet.js, xlsx.js    Builds the .xlsx (no dependencies)
test/                          Tests: the real Code.js running on in-memory fakes of the Google services
```

```bash
npm test             # runs the tests (they also check the .xlsx is up to date)
npm run build        # rebuilds template/parts-inventory.xlsx after changing Code.js
```

`node dev/build-sheet.js --base64` prints the .xlsx for uploading through the Google Drive API, which
converts it to a Google Sheet. Google's import has two limits the builder works around: it can't make
tick boxes (Ordered is a *Yes* dropdown until setup turns it into tick boxes) and it rejects
open-ended ranges like `A2:A`, so the formulas use whole columns (`A:A`).

**Deploying with [clasp](https://github.com/google/clasp):** enable the Apps Script API at
[script.google.com/home/usersettings](https://script.google.com/home/usersettings), run `npx @google/clasp login`,
then copy `.clasp.json.example` to `.clasp.json` with the script ID of the spreadsheet's project
(Apps Script → Project Settings) and run `npm run push`.

**Notes on the design**

- Columns are found by their headings (`FIELDS` in `Code.js`), so the sheet can be rearranged.
- Everything that reads and writes the sheet runs under a script lock, so two edits at once can't send
  the same email twice.
- The `@OnlyCurrentDoc` annotation in `Code.js` makes Google ask for access to this one spreadsheet
  rather than all of them, which works because the script only uses `getActiveSpreadsheet()`. The tests'
  fakes have no `openById`, so they fail if that changes.
