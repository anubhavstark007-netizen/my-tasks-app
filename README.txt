MY TASKS — SHARED TEAM VERSION
================================

This package contains:
- index.html: updated website interface
- Code.gs: Google Apps Script backend
- manifest.json, sw.js, icon.svg: PWA files

IMPORTANT
---------
Tasks are saved in a Google Sheet. They are NOT erased after 24 hours or when the date changes.
The "Completed Today" screen shows tasks completed today. Use "View all completed tasks" there to see older completed records. A task is permanently removed only when someone presses the delete (×) button and confirms.

The name picker remembers the selected person for the current date in that browser. It asks again on the next calendar day. This is a simple person selector, NOT secure login. Anyone with the website URL can view/edit/delete shared tasks.

PART A — Create the shared Google Sheet/backend
------------------------------------------------
1. Open Google Sheets and create a NEW blank spreadsheet. Name it: My Tasks Database.
2. In that spreadsheet, click Extensions > Apps Script.
3. Remove any starter code in Code.gs and paste the full contents of the Code.gs file from this package.
4. Click Save.
5. In the function dropdown beside Run, choose setupDatabase, then click Run.
6. Approve Google permission prompts. After it finishes, return to the spreadsheet; a tab named "Tasks" should exist.
7. In Apps Script, click Deploy > New deployment.
8. Click the gear/type selector and choose Web app.
9. Set "Execute as" to Me.
10. Set "Who has access" to Anyone (or the broadest option allowed for your Google account).
11. Click Deploy and authorize if asked.
12. Copy the Web app URL ending in /exec. Treat this URL as public because it is embedded in the website.

PART B — Connect the website to the Sheet
-----------------------------------------
1. Open index.html in a text editor (Notepad is fine).
2. Near the start of the JavaScript, find:
   const API_URL = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";
3. Replace ONLY the placeholder text with your copied /exec URL, keeping the quotation marks. Example format:
   const API_URL = "https://script.google.com/macros/s/DEPLOYMENT_ID/exec";
4. Save index.html.

PART C — Upload updated website files to GitHub
------------------------------------------------
1. Open your existing GitHub repository: my-tasks-app.
2. Upload/replace these four files in the repository root:
   - index.html
   - manifest.json
   - sw.js
   - icon.svg
3. Commit the changes.
4. Open the GitHub Pages website and refresh it. If the old version remains, close the installed app/browser tab and reopen; the service worker cache version has been updated.
5. Test by adding a sample task, then open the Google Sheet and check the Tasks tab. Open the website in another browser/device to confirm the same task appears.

HOW TO USE
----------
- Each calendar day, the site asks who opened it: Aakash, Anuj, Rishabh, Anubhav, or Surbhi.
- Home shows the three shared follow-up lists and all five members.
- Open a member to add/view that person's tasks.
- Open a follow-up list to add/view that category's tasks.
- Tick a checkbox to mark a task completed. It remains in the Sheet and in completed history.
- Completed Today shows tasks completed on the current date. "View all completed tasks" shows older completed tasks too.
- Press × and confirm only when you really want to permanently delete a record.

TROUBLESHOOTING
----------------
- If it says API_URL placeholder: the /exec URL was not pasted into index.html.
- If the page says it cannot connect: check that deployment is a Web app, access is set to Anyone, and the URL ends with /exec.
- If you change Code.gs after deploying: Apps Script > Deploy > Manage deployments > Edit (pencil) > choose New version > Deploy.
- Do not run setupDatabase again after you start using the app: it clears the Tasks tab and would erase its records.
