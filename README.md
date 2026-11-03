# Project Teams Ops Hub

The website Cornell Project Teams staff use for safety and operations: Space Status reports,
projects, the tool inventory, team member photos, Fabman, compliance and the fleet.

Live site: https://cornellprojectteams.github.io/centralhub/

## Guides

Each topic is covered in one place only.

| If you want to | Read |
|----------------|------|
| Report an issue or finish a job | [Staff how-to](space-status-how-to.html) |
| Do the everyday admin jobs | [Admin guide](docs.html) |
| Change the site or a script, set up a web app, or put a change live | [Developer guide](dev-guide.html) |
| Try things on fake data | [Sandbox](sandbox/README.md) (start it with `node sandbox/server.js`) |

The guides are pages on the site, so they open from the admin and developer pages too.

## Quick start for developers

```bash
node sandbox/server.js
```

Open http://localhost:8080/sandbox/. Edit, save, refresh. When it works:

```bash
git add -A
git commit -m "what changed"
git push
```

The live site updates within a minute or two. Scripts in `apps-script/` don't go live from a push;
the [developer guide](dev-guide.html#scripts) says how.
