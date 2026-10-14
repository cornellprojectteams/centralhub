# Project Teams Ops Hub

Static web hub for Cornell Project Teams Safety & Operations. Links staff and admins to everyday Google Forms, Sheets, Docs, and Apps Script tools from one page.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Staff page: everyday tasks + the Operations Team open-tasks panel |
| `admin.html` | Admin page: back-office tools. Reached by shared link, not in the staff nav |
| `command-center.html` | Ops metrics dashboard. Reached by shared link |
| `team.html` | Shareable per-team open-issues page (`team.html?team=<name>`) |
| `styles.css` | Layout and Cornell theming |
| `app.js` | Live search, open-issues dashboard panel, admin sidebar (shared by the pages) |
| `config.js` | Single source for the Space Status web app `/exec` URL and the inventory settings (`INVENTORY`) |
| `inventory.html`, `inventory.js`, `inventory.css` | The tool & equipment inventory app (also opens in a panel from the staff and admin pages) |
| `apps-script/08_inventory.gs` | The inventory's backend: builds the Google Sheet, keeps it tidy, and serves the app's API |
| `assets/` | Cornell seal and favicon SVGs |

## Running locally

Any static file server works, for example:

```bash
npx serve .
```

Open `http://localhost:3000` (or the port shown).

## Navigation

- Header buttons switch between **Staff** and **Admin tools**.
- Only one hub zone is visible at a time.
- `#staff` and `#admin` URL hashes restore the selected view on load or refresh.

## Search

The search box filters tools by visible text and optional `data-keywords` attributes on categories, actions, and footnotes. While searching, matching items from both zones may appear. Clearing search returns to the active nav view.

## Adding a tool link

1. Choose the correct zone in `index.html` (`#staff` or `#admin`).
2. Add an `<a class="action">` with `href`, `target="_blank"`, and `rel="noopener"`.
3. Include `<span class="action-title">` and `<span class="action-sub">` for the label and subtitle.
4. Add `data-keywords="..."` with abbreviations and synonyms users might search for.

Admin tools are grouped in `<section class="category admin-block">` blocks with an `<h3 class="admin-label">` heading.

## Styling notes

- Staff actions use the red gradient theme (`--red-*` tokens).
- Admin actions use the burgundy gradient theme (`--burgundy-*` tokens).
- Design tokens live in `:root` at the top of `styles.css`.

## Completion evidence &amp; Projects

Backed by `apps-script/04_tasks_projects.gs`, served by the same Space Status web app.

**Tasks are the open issues; there is no separate task list.** Completing any open
action item (in the Open-issues dashboard, a team portal, or via the *Mark complete*
button in the notification email) now works like this:

1. The doer uploads a photo as **evidence of completion** (required; the single
   upload *is* the completion action).
2. Submitting fires a client-side **confetti burst** and moves the item to
   *Pending approval* (stamps `Completed at`; reminder emails pause).
3. An admin **Approves** → stamps `Addressed at` (= Completed/resolved), or
   **Sends back** → clears the submission (back to Open / Uncompleted).

This reuses the one `Form Responses` tracking sheet, adding two columns
(`Completion photo`, `Completed at`), with no parallel system.

**Projects** (`?module=projects`, linked from the hub via a `data-module` attribute
that `app.js` rewrites to `…/exec?module=projects`) are multi-user work
created/assigned by an admin. Any student can **pick (join)** a project; the first
pick auto-flips it *Assigned → In Progress*. Multiple assignees are tracked.
Completion is a manual trigger that **enforces both a before and an after photo**.

**Roles.** Everyone is a doer by default. Admin-only actions (approve or send back
completions, create/assign projects) are gated by the shared passcode
(`CONFIG.toolPassHash`, same as the other admin tools), validated server-side on
every mutating call; unlock with the **Admin** bar on the page. Project completion
is open to assignees. Photo uploads go to a dedicated Drive folder.

**First run:** open the Space Status Apps Script project, run `setupTasksProjects()`
once (creates the `Projects` tab and the two completion columns), then push a **new
deployment version** (any change to web-app logic needs one); the URL stays the same.

## Tool & equipment inventory

The inventory is an app inside the hub (`inventory.html`). A Google Sheet is its database, and
that sheet's own script (`apps-script/08_inventory.gs`) is both its caretaker and its API. It
replaced the old equipment registry (the `03` and `06` scripts) in October 2026. Tools stay in the
shops and nobody borrows them, so the inventory records what there is, where it lives, what shape it
is in, its maintenance, and how much of each supply is left.

- **Staff page > Tool inventory** opens it in a panel. The board works like a shop shadow board:
  each tool sits on a steel pegboard, hazard tape marks anything out of service, a hang tag flags a
  repair or a supply running low, and retired items fade. Search, the status chips and the Team,
  Location and Category dropdowns narrow it down. A tool's page shows where it lives, its condition,
  its details and its maintenance.
- **Admin page > Equipment & inventory** opens the admin view: add and edit items, photos, teams,
  maintenance, stock counts and test data, plus printable QR labels and the spreadsheet itself.
- A tool's QR label opens that tool (`inventory.html?item=PT-001`). The browser's Back button closes
  an open tool.

**Photos.** Admins add a photo from a tool's page or in the add/edit form (on a phone, it can be
taken on the spot). The app shrinks it before sending, and the script saves it in an *Inventory
photos* folder next to the sheet in Drive, shared view-only by link like the Space Status uploads.
The Photo URL column links to it. Replacing or removing a photo moves the old file to the trash.

**Teams.** Every item has a Team: the project team that owns it, or *Shared* for the program's own
tools. The team list (the official team names) is on the Lists tab, column E. A new location or
subcategory typed in the app's add/edit form is added to the Lists tab as well, so the sheet's
dropdowns offer it too.

The app talks to the sheet JSONP-style, like the Command Center, so the static site needs no CORS
setup; photos are the one exception and go up as a plain POST. The app keeps the last good copy in
the browser, so it paints at once and still shows something if the sheet is slow. The admin view
follows the hub's convention: it opens from the unlisted admin page (`?admin=1`) with no passcode.

| Tab | What it holds |
|-----|---------------|
| Dashboard | The same numbers as the app, for anyone working in the sheet |
| Inventory Master | One row per tool or piece of equipment, plus consumables (Current Qty / Min Qty) |
| Maintenance Log | Repairs, calibrations, inspections. Next Due feeds the app and the Dashboard |
| Lists | Dropdown options (categories with their ID prefix, subcategories, locations, teams) and links |

Status is *In service* unless Out of Service says Maintenance or Retired; consumables show In
Stock, Low Stock or Out of Stock. Item IDs (PT-001, EQ-004, ...) are assigned from the category's
prefix and never change, so they are safe to print on labels.

**Connecting it** (once)

1. In the inventory Google Sheet: Extensions > Apps Script, replace the code with
   `apps-script/08_inventory.gs`, save, and run `setupInventory` (safe to re-run; it repairs and
   runs a self-test). Approve the access it asks for, including Drive (for photos). Never run it in
   the old *Tool inventory, purchasing and locks* sheet: its Mill room Tooling and Locks tabs feed
   the Command Center, and the script refuses to run there.
2. Deploy > New deployment > type **Web app**. Execute as: **Me**. Who has access: **Anyone**.
   Copy the URL that ends in `/exec` into `config.js` as `INVENTORY.apiUrl`.
3. After any later change to the script: Deploy > Manage deployments > pencil > Version: New
   version > Deploy. The URL stays the same.

**Updating a sheet set up when the inventory had check outs.** Paste in the new `08_inventory.gs`
and run `setupInventory` once. It unlinks the backup Google Form and moves it to the Drive trash,
deletes the Checkout Log tab (if it holds anything besides test data, it keeps it as *Old checkout
log* instead), removes the Current Holder and Due Date columns, and renames Checkout Link to Item
Link. Items, maintenance and photos stay as they are. Then deploy a new version as in step 3.

**Updating a sheet set up before photos and teams.** The same `setupInventory` run asks for Drive
access, moves the new Team column next to Location and fills in the team list. Until then the app
keeps working and the admin view says photos and teams are switched off.

**Test data.** The admin view's *Load test data* button (or Inventory menu > Add test data in the
sheet) adds 26 sample items with teams and maintenance. Everything it adds is tagged
`[test data]`. *Reset test data* swaps in a fresh copy (its dates are relative to today), and
*Remove test data* takes all of it out again, including photos added to test items, while leaving
real items alone.

**Retiring the old registry.** Its code is gone from this repo; these steps clean up the Google side:

1. In the *Tool inventory, purchasing and locks* sheet, open Extensions > Apps Script > Triggers and
   delete the `dailyCmmsCheck` trigger if it is there, then delete the registry code file.
2. In that sheet, delete only the registry tabs: Equipment, Inventory, Maintenance, Action items,
   Checkouts. Keep Team tool inventory, Mill room Tooling, Facility non recurring and Locks.
3. In the Space Status Apps Script project, delete the `06_registry` file, replace
   `02_notify_on_submit` with this repo's version, then Deploy > Manage deployments > New version.
   Old registry links and printed QR labels then land on an "Inventory has moved" page.

## Related docs

`SHEETS_INVENTORY.md` describes the underlying Google Workspace assets this hub links to.
