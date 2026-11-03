# Oceanside Promo Planner

Answers one question first: what do we say this Sunday? Built from the Mission Control playbooks.

- **To do**: what's overdue, what's about to start, every task by day, and where each running campaign is
- **Timeline**: one row per campaign, one column per week. Rows open to show each week's tasks
- **Campaign**: the whole plan laid out week by week, like the playbook
- **Sunday**: the lineup (1 stage, up to 2 Booth) from the campaigns due an announcement
- **Rules**: what each level gets, and a reply for when someone asks for more
- **`?ask`**: public questionnaire for ministry leaders. It suggests a level and sends a request

## Stack

`index.html` on GitHub Pages, a Google Sheet behind an Apps Script web app (`apps-script/Code.gs`). No build step.
Reading and editing need the office code. The public can only send a request.

## Setup

1. New Google Sheet, Extensions → Apps Script, paste `apps-script/Code.gs`, save.
2. Run `setup`. The office code is in the Execution log. Change it in Project Settings → Script properties → `OFFICE_CODE`.
3. Optional: add a `NOTIFY_EMAIL` script property to get an email for each request.
4. Deploy → New deployment → Web app. Execute as Me, access Anyone. Copy the URL.
5. Put the URL in `CONFIG.apiUrl` in `index.html` and push.

After changing `Code.gs`: Deploy → Manage deployments → pencil → New version.

## How the lineup is chosen

For each Sunday, campaigns whose playbook week calls for an announcement compete by level, then by nearest date.
Level 1 or 2 takes the stage, the next two take the Booth, and anything left is bumped to the bulletin.
Levels 4 and 5 never enter automatically. An always-on next step only fills a Sunday with fewer than two announcements.
Anyone can change a Sunday by hand, but there are only three slots.

## Change the rules

Playbooks, level perks and policy answers are the `PB`, `LEVEL` and `ASKS` tables near the top of the script in `index.html`.

## Local preview

Serve the folder on localhost with no `apiUrl` and it runs on `_local/demo.js` (not committed). Add `?reset` to start over.
