# Team member photo database

`team-photos.html` shows every project team's public member photos, with filters for team and
role. The data is one file, `data/team-photos.json`. This folder holds the scripts that fill it
from the team websites.

| File | Purpose |
|------|---------|
| `sites.json` | Each team's website and team page. Shown on the photos page as "Team site" |
| `targets.json` | The pages the scraper opens for each team. Edit this when a team moves its team page |
| `scrape.py` | Opens each target in a headless browser and saves every photo card to `raw.json` |
| `extract.js` | The in-page code `scrape.py` runs to find photos and the text next to them |
| `build.py` | Reads `raw.json`, works out name, role, lead or member, subteam and class year, and merges the result into `data/team-photos.json` |
| `raw.json` | The last scrape, kept so `build.py` can be rerun without scraping again |

## Refresh from the team sites

Once a semester, or when teams post new members:

```bash
pip install playwright && playwright install chromium
python3 tools/team-photos/scrape.py
python3 tools/team-photos/build.py --dry-run
python3 tools/team-photos/build.py
```

`--dry-run` prints how many people are new, changed and no longer listed without writing anything.
To scrape only some teams, name them: `python3 tools/team-photos/scrape.py "Mars Rover" Baja`.
Then commit `data/team-photos.json` and push. The page updates when GitHub Pages redeploys.

What a refresh does:

- A person on a team site is added, or their role, photo, subteam and class year are updated.
- A person who is no longer on their team's site is marked `"status": "left"`. The page hides
  them unless *Show graduated and former* is ticked.
- A team whose site fails to load is left as it was. Nobody on it is marked left.

## Remove graduates

Nothing has to be done each May. The page treats anyone whose class year has passed as
graduated (from June 1 of that year) and hides them by default.

To mark someone graduated who has no class year on file, or graduated early, set their
`status` to `"graduated"` in `data/team-photos.json`. A later refresh keeps it.

To delete everyone who has graduated or left from the file:

```bash
python3 tools/team-photos/build.py --prune
```

## Add or correct a person by hand

Edit `data/team-photos.json` (GitHub's web editor works). Each person is one entry in `people`:

```json
{
 "id": "manual-jdoe1",
 "team": "Baja",
 "name": "Jane Doe",
 "role": "Suspension Lead",
 "roleType": "lead",
 "subteam": "Suspension",
 "classYear": 2028,
 "photo": "https://baja.mae.cornell.edu/assets/img/headshots/Jane_Doe.webp",
 "sourcePage": "",
 "status": "current",
 "source": "manual"
}
```

- `team` must match a name in `sites.json`.
- `roleType` is `lead`, `member` or `advisor`. The page's role filter uses it.
- `photo` is an image URL. Leave it `""` and the page shows initials.
- `"source": "manual"` means a refresh never changes or removes the entry. Give it any unique `id`.
- To fix a scraped person so a refresh does not overwrite the fix, edit the entry and add
  `"locked": true`.

## When a team's page changes

If a refresh shows a team with far fewer people, open its team page and check the address.
Update the team's entry in `targets.json` (and `teamPage` in `sites.json`), then rerun both scripts.
For a single-page site where the team view opens from a nav link, write the target as
`"https://site/|Link text"`.

Hack4Impact has no public team page with member photos, so it has no targets.
