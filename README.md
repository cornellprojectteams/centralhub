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
| `config.js` | Single source for the Space Status web app `/exec` URL |
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

**Tasks are the open issues — there is no separate task list.** Completing any open
action item (in the Open-issues dashboard, a team portal, or via the *Mark complete*
button in the notification email) now works like this:

1. The doer uploads a photo as **evidence of completion** (required — the single
   upload *is* the completion action).
2. Submitting fires a client-side **confetti burst** and moves the item to
   *Pending approval* (stamps `Completed at`; reminder emails pause).
3. An admin **Approves** → stamps `Addressed at` (= Completed/resolved), or
   **Sends back** → clears the submission (back to Open / Uncompleted).

This reuses the one `Form Responses` tracking sheet, adding two columns
(`Completion photo`, `Completed at`) — no parallel system.

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
deployment version** — any change to web-app logic needs one; the URL stays the same.

## Related docs

`SHEETS_INVENTORY.md` describes the underlying Google Workspace assets this hub links to.
