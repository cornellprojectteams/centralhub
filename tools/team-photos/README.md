# Team photo scraper

The scripts behind `team-photos.html`. Everything about them (the Update button, running them by
hand, editing a person, fixing a team whose page moved) is in the developer guide:
[dev-guide.html#photos](../../dev-guide.html#photos).

| File | Purpose |
|------|---------|
| `sites.json` | Each team's website and team page |
| `targets.json` | The pages the scraper opens for each team |
| `scrape.py` | Opens each target in a headless browser and saves the photo cards to `raw.json` |
| `extract.js` | The in-page code `scrape.py` runs |
| `build.py` | Turns `raw.json` into `data/team-photos.json` |
| `raw.json` | The last scrape |
