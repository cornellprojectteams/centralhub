# Trying a change on real Google

The local sandbox is the place to build and test. Before a script change goes into a live Apps Script project, it is worth one run on real Google: real Gmail, real Drive photos, a real web app deployment. This kit makes copies of the Space Status scripts that are safe to run in your own Google account.

The copies:

- use a spreadsheet and a form of your own, which `sandboxSetup` makes for you
- send every email to you only, with [SANDBOX] in front of the subject and a yellow note at the top saying who would have received it in production
- point the email buttons at your own copy of the web app

Nothing in them can open the real Space Status sheet or email a team. `make-copies.js` checks this every time and refuses to write the files if a real sheet id or a real address is still in them.

## 1. Make the copies

In the repo folder:

```
node sandbox/google/make-copies.js you@cornell.edu
```

Use the address of the Google account you will test with. The files land in `sandbox/google/out/space-status/`, which git ignores.

## 2. Make an Apps Script project of your own

1. Go to script.google.com, signed in as yourself, and press New project. Rename it "Space Status (sandbox copy)".
2. Open Project Settings (the gear) and set the time zone to Eastern Time (New York), like the live project.
3. Back in the Editor, add a script file for each of the four files in `out/space-status/` (the plus next to Files, then Script). Name each one like the file, without .gs, and paste the file's contents. Delete the empty Code.gs.

## 3. Set it up

1. In the function menu above the code, choose `sandboxSetup` and press Run.
2. Google asks for permission the first time. It is your own project in your own account; allow it.
3. The log at the bottom lists your new sheet, your new form and what to do next.

`sandboxSetup` makes "Space Status (sandbox copy)" in your Drive with the same tabs as the real sheet, a "Space Status Report (sandbox copy)" form linked to it, a Team Contacts tab where every team's address is yours, and the form-submit trigger that sends the emails.

Apps Script cannot add a file-upload question to a form. If you want to test report photos, open the form and add a File upload question titled exactly `Photo Upload`.

## 4. Publish your copy of the web app

1. Deploy, New deployment. Under Select type (the gear), choose Web app.
2. Execute as: Me. Who has access: Only myself.
3. Deploy, and keep the web app URL.

The buttons in the emails find this address on their own. If a button opens the wrong page, put the URL in a script property named `SANDBOX_WEBAPP_URL` (Project Settings, Script properties).

## 5. Try it

- Submit a report on your form, or run `sandboxSubmitTestReport`. The email arrives in your inbox.
- Press Mark complete in the email and finish the job with a photo.
- Open your web app URL with `?view=all&admin=1` at the end to see the admin dashboard and approve it. `?module=projects&admin=1` and `?module=proposals` open projects and proposals.
- `sandboxStatus` prints all the links again.
- `sandboxDailyReminders` turns on the morning reminder job, like the live project has. Skip it unless you are testing reminders.

## 6. Test a change

1. Make the change in the repo (`apps-script/02_notify_on_submit.gs` and friends) and try it in the local sandbox first.
2. Run `make-copies.js` again and paste the changed file over its copy in your project.
3. Deploy, Manage deployments, the pencil icon, Version: New version, Deploy. Or use Deploy, Test deployments, whose URL always runs the latest saved code.

When it works, the change goes into the live project from the original file in `apps-script/`, never from a copy in `sandbox/google/out/`. The copies say so in their first line.

## 7. Clean up

Run `sandboxRemoveTriggers`, then delete the sheet, the form and the project from your Drive whenever you like.

## The other scripts

- **Inventory (08_inventory.gs)** is safe to try on Google as it is. Make a new blank Google Sheet, open Extensions, Apps Script, paste the file, run `setupInventory`, then use the Inventory menu's Load test data. It works only with the sheet it is pasted into.
- **Fabman (05_fabman_members.gs)** must not be copied. It holds the live Fabman key and changes real member accounts. Use the local sandbox's fake Fabman.
- **Command Center (command-center/Code.gs)** reads the real sheets. It only reads, but that is real people's data, so test it in the local sandbox.
- **Fleet driver cleanup** changes the sheet it is bound to. Test it in the local sandbox, where the drivers sheet is made up.
