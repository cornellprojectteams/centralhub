# The Ops Hub sandbox

A full copy of the platform that runs on your own computer with made-up data. The staff page, the admin page, the Space Status web app, the inventory, the Command Center and the sheet scripts all work, so you can click through every flow, change any file and see what happens. Nothing you do here reaches Google, Fabman or a real person.

## Start it

**In a Codespace (nothing to install).** On the repo's GitHub page, open the green Code button, choose the Codespaces tab and create a codespace. The sandbox starts by itself. When the "Ops Hub sandbox" port pops up, open it in the browser. If you closed that tab, open the Ports panel and click the globe next to port 8080.

**On a laptop.** Install Node.js (version 18 or newer) from nodejs.org, then in a terminal, inside the repo folder:

```
node sandbox/server.js
```

Open http://localhost:8080/sandbox/ in your browser. Stop it with Ctrl+C. There is no npm install step; the sandbox only uses what comes with Node.

Other options:

- `node sandbox/server.js --fresh` starts from the original fake data instead of where you left off.
- `node sandbox/server.js --port 9000` uses another port.
- `node sandbox/server.js --host 0.0.0.0` lets a phone on the same Wi-Fi open it, to try the pages on a small screen.

## What is real and what is fake

Real, straight from the repo, so your edits show up on refresh:

- Every page of the site (index.html, admin.html, inventory.html, command-center.html and the rest), with its CSS and JavaScript.
- Every Apps Script file in `apps-script/`. The sandbox runs them itself, the way Google would.

Fake, and safe to break:

- **Sheets.** Each spreadsheet the platform uses has a made-up copy with the same id, the same tabs and the same headers. Open them from Sheets in the top bar. Click a cell to change it.
- **Gmail.** Every email a script sends goes to the Outbox. Nothing is delivered. You will see the real always-CC address on some emails; that is just what the script asked for.
- **Drive.** Photos uploaded through the forms and the apps are kept inside the sandbox.
- **Forms.** The Space Status report form, the ELL shift form and the FabMan Agreement have stand-ins under "Make something happen" on the control panel. Submitting one writes a row and runs the same trigger Google would. The FabMan Agreement stand-in behaves like the Qualtrics survey: its workflow adds a row to the Fabman sheet, which sets off the sheet's change trigger.
- **Fabman.** The Fabman script talks to a fake Fabman with 42 made-up members. See what it changed on the Fake Fabman page, where you can also make Fabman fail or refuse the API key to see how the scripts cope.
- **Time.** The daily reminder job does not run on a clock here. Use "Run today's reminder job", or "Skip ahead a day" to make every date in the sandbox one day older first.

All people, NetIDs and email addresses are invented, and every email address ends in sandbox.test, a domain that cannot receive mail. The one exception is the fake Fabman members, whose addresses follow the NetID@cornell.edu pattern the Fabman script looks people up by; nothing is ever sent to them. Team names are the real ones so the pages look familiar.

## Staying safe

- Links that go to a real Google Doc, Sheet, Form, Drive file, Apps Script project or Fabman open a "Leaving the sandbox" page first. Reading a real doc is fine. Changing a real sheet or submitting a real form is not a test.
- Links to the platform's own sheets, forms and web apps are swapped for the sandbox copies, so the admin page's buttons stay inside the sandbox.
- Scripts cannot reach the internet. The only outside address they can call is the Fabman API, and that call goes to the fake Fabman.
- The sandbox never serves the real data exports (the downloads and email-notifications folders), and it never uses the Fabman key in `apps-script/05_fabman_members.gs`.
- The sandbox's data lives in `sandbox/.state/`, which git ignores. Your fake data never ends up in a commit.

## Working in it

The control panel (http://localhost:8080/sandbox/) is home base. From there:

- **Staff page and Admin page** are the real hub pages. The small "Sandbox" tab in the corner brings you back to the control panel.
- **Outbox** lists every email, newest first. Open one to see it the way the team would. Its buttons work and open the sandbox's pages.
- **Sheets** shows the fake spreadsheets. Click a cell to edit it (Enter saves, Esc cancels, Tab moves right). Edits run the sheet's edit triggers, like on Google. Menus that a script adds when the sheet opens (Driver cleanup, Inventory) appear above the grid, and the Fabman sheet's buttons appear as chips on the tab they belong to.
- **Scripts** lists the Apps Script projects. On a project's page you can run any function, see what it logged, which triggers it installed and its script properties. When a script shows a dialog, the "If a script asks" picker decides whether the sandbox answers Yes or No.
- **Reset the sandbox** (bottom of the control panel) puts every sheet, photo, email and fake Fabman member back the way they started. It never touches your code.

To change something, edit the file and refresh the page. That goes for the pages, the CSS, the JavaScript and the `.gs` files. There is no build step and no restart.

## Things to try

The control panel has a checklist that walks through the whole platform. In short:

1. File a report on the Space Status form, then find the team's email in the Outbox.
2. Press the button in that email and finish the job with a photo. Approve it on the Admin page, in the Open issues dashboard.
3. Send another job back with a reason and read the email the team gets.
4. Skip ahead a day and see which reminders go out.
5. Create a project on the Admin page, join it from the staff page, finish it and approve it.
6. Rate three proposals with a made-up NetID and watch the cord reels proposal turn into a project.
7. Add a tool in the inventory, print its QR label and log maintenance on it.
8. Open the Command Center and see your changes there.
9. Run Driver cleanup on the drivers sheet, then Flag graduates and Process graduates in the Fabman sheet.
10. Seven FabMan Agreement sign-ups are waiting without accounts. Use the Fabman sign-ups menu in the Fabman sheet to check them and catch them up, then sign the agreement yourself and watch the Fabman status column.
11. Change a color in styles.css or an email subject in `apps-script/02_notify_on_submit.gs`, and refresh.
12. Break a script on purpose, see how the error is reported, then fix it.

## When something breaks

- A broken `.gs` file shows an Error page with the file and line, the same way Apps Script would. The control panel also lists recent script errors.
- Each project's page under Scripts shows its recent log, including anything the script printed with `Logger.log` or `console.log`.
- `node sandbox/check.js` runs a quick self-test of the whole sandbox (no browser needed) and says what failed.
- If the fake data gets into a state you do not like, press Reset or start with `--fresh`.
- If port 8080 is busy, the sandbox is probably already running in another terminal. Use that one, or pass `--port`.

## What the sandbox does not do

It is a stand-in, not a copy of Google, so a few things behave differently:

- Sheet formulas are shown but not calculated, except the inventory's Status and Next Maintenance columns, which the sandbox works out the way the sheet's formulas would.
- Triggers run when the matching thing happens here (a form submit, a cell edit, a sheet opening). Time-driven triggers do not run on their own; use the reminder buttons.
- Sharing settings on Drive files are recorded but not enforced, and there are no quotas or execution limits apart from a one-minute cap per run.
- Everything runs on New York time, like the real projects.
- Only the Apps Script features the platform uses are there. If you use one that is missing, the error says so. Add it in `sandbox/lib/gas.js`.

So a change that works here still needs a real test on Google before it goes live. See "Trying a change on real Google" below.

## From the sandbox to the live platform

The site and the scripts go live in different ways.

- **Site files** (HTML, CSS, JavaScript, config.js) go live when they are pushed to the main branch on GitHub. GitHub Pages publishes them within a minute or two.
- **Apps Script files** do not go live from the repo. Someone with access pastes the changed file into the matching Apps Script project. For the web apps (Space Status, inventory, Command Center), they then publish a new version: Deploy, Manage deployments, the pencil icon, Version: New version, Deploy. The URL stays the same.

The Developer guide (dev-guide.html) has the details, including which Apps Script project each file belongs to.

### Trying a change on real Google

Before pasting a script change into a live project, try it in your own Google account. `sandbox/google/` has a kit that makes safe copies of the Space Status scripts: they use a sheet of your own, and every email goes only to you. Its README walks through it.

## How the sandbox is built

Everything is in the `sandbox` folder and uses plain Node with no packages.

- `server.js` serves the site and the sandbox pages, and swaps real addresses for sandbox ones on the way out.
- `lib/gas.js` is the stand-in for Apps Script: SpreadsheetApp, DriveApp, GmailApp, HtmlService, the cache, properties, triggers, UrlFetchApp and the rest.
- `lib/projects.js` lists the Apps Script projects (which files, which sheet they are bound to) and runs their functions. Every run starts fresh, like on Google.
- `lib/fixtures.js` builds the fake data. Change it to start from different data, then press Reset.
- `lib/forms.js` has the stand-in forms and their questions.
- `lib/rewrite.js` and `public/guard.js` keep the pages inside the sandbox.
- `lib/pages.js`, `public/sandbox.css` and `public/sandbox.js` are the sandbox's own pages.
- `lib/store.js` saves the fake data to `sandbox/.state/state.json`.
- `check.js` is the self-test.

If the platform gets a new spreadsheet or a new Apps Script project, add it to `lib/projects.js` and give it fake data in `lib/fixtures.js`, and the sandbox will run it like the others.

The sandbox reads the web app URLs and the sheet ids from `config.js`, `app.js` and the scripts when it starts. After changing one of those, restart it with `--fresh` so the fake data follows. Changes to the sandbox's own files (anything in `sandbox/`) also need a restart.
