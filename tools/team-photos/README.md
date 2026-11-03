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
| `../../.github/workflows/team-photos.yml` | The GitHub Action the Update button starts |
| `../../apps-script/09_team_photos.gs` | The web app between the button and GitHub. It holds the GitHub token |

## Refresh from the team sites

Open the photos page from the admin page (Teams > Team member photos) and press
**Update from team sites**. The page shows each step and reports what changed when it is done,
usually 5 to 10 minutes later. It also runs on its own on the 1st of each month.

How it works: the button asks the update web app (`apps-script/09_team_photos.gs`) to start the
GitHub Action `.github/workflows/team-photos.yml`. The Action runs `scrape.py` and `build.py` and
commits the new `data/team-photos.json`. The page follows the run, waits for GitHub Pages to
publish the new file, then reloads the photos. A run in progress is remembered, so closing the
page does not stop it, and reopening the page picks it up again.

What a refresh does:

- A person on a team site is added, or their role, photo, subteam and class year are updated.
- A person who is no longer on their team's site is marked `"status": "left"`. The page hides
  them unless the *Graduated & former* switch is on.
- A team whose site fails to load keeps its last good scrape. Nobody on it is marked left, and
  the page names the team when the update finishes.
- A team whose site suddenly lists less than half as many people is treated the same way. This
  usually means the team moved its team page (see *When a team's page changes*).

| What the page says | What to do |
|--------------------|------------|
| The update service isn't connected yet | Do *Connect the Update button* below |
| Couldn't reach / didn't answer | Try again. If it keeps happening, check `TEAM_PHOTOS.refreshUrl` in `config.js` |
| The update couldn't start: token | Replace `GITHUB_TOKEN` in the web app's script properties |
| Already up to date | An update finished in the last 5 minutes. Wait and try again |
| GitHub didn't start the update | Check that Actions are on: repository Settings > Actions |
| The update failed on GitHub | Open the run log from the link. The data was not changed |

### Connect the Update button

Once, by someone with admin rights on the GitHub organization:

1. On GitHub: Settings > Developer settings > Fine-grained tokens > Generate new token.
   Resource owner: `cornellprojectteams`. Repository access: only `centralhub`.
   Permissions: **Actions: Read and write**. Copy the token.
2. Go to script.google.com > New project. Paste in `apps-script/09_team_photos.gs`.
3. Project Settings > Script properties > Add `GITHUB_TOKEN` with the token as its value.
4. Deploy > New deployment > type **Web app**. Execute as: **Me**. Who has access: **Anyone**.
   Copy the URL that ends in `/exec`.
5. In `config.js`, set `TEAM_PHOTOS.refreshUrl` to that URL. Commit and push.

The token expires on the date you chose in step 1. When it does, the page says so; make a new one
and replace the script property. No redeploy is needed for that.

### Run it by hand

To run the scraper on your own computer instead (for example to test a change to `targets.json`):

```bash
pip install playwright && playwright install chromium
python3 tools/team-photos/scrape.py
python3 tools/team-photos/build.py --dry-run
python3 tools/team-photos/build.py
```

`--dry-run` prints how many people are new, changed and no longer listed without writing anything.
To scrape only some teams, name them: `python3 tools/team-photos/scrape.py "Mars Rover" Baja`.
Then commit `data/team-photos.json` and push.

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
