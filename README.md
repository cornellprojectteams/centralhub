# Project Teams Ops Hub

Static web hub for Cornell Project Teams Safety & Operations. Links staff and admins to everyday Google Forms, Sheets, Docs, and Apps Script tools from one page.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Main hub: staff tasks, admin tools, search, navigation |
| `styles.css` | Layout, Cornell-themed staff (red) and admin (burgundy) styling |
| `app.js` | View switching (`#staff` / `#admin`) and live search filtering |
| `admin.html` | Redirects to `index.html#admin` for bookmark compatibility |
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

## Related docs

`SHEETS_INVENTORY.md` describes the underlying Google Workspace assets this hub links to.
