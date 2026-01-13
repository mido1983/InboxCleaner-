# InboxCleaner

InboxCleaner is a Gmail sidebar add-on that helps you bulk trash or archive messages that match a phrase or a preset query. It uses Google Apps Script CardService UI and the Advanced Gmail Service with the minimum Gmail scope for modifying messages.

## Assumptions
- Preset "Clean now" actions use Trash mode and bypass dry-run.
- Large runs are capped at 2000 messages per execution. Re-run the same query to continue.
- Progress is summarized after the run because Gmail add-ons do not support live progress updates.

## Prerequisites
- Node.js 18 or later
- `clasp` available via `npx` using the project dev dependency

## Setup
1) Install dependencies

```
npm install
```

2) Authenticate clasp

```
npx clasp login
```

3) Create the Apps Script project

```
npx clasp create --type addon --title "InboxCleaner" --rootDir ./src
```

4) Push the local files

```
npx clasp push
```

5) Open the Apps Script project

```
npx clasp open
```

## Enable Advanced Gmail Service
1) In the Apps Script editor, open Services.
2) Click the plus button.
3) Select Gmail API and click Add.

## OAuth scopes
The scopes are defined in `src/appsscript.json`:
- `https://www.googleapis.com/auth/gmail.addons.execute`
- `https://www.googleapis.com/auth/gmail.modify`

## Deploy as a Gmail add-on
1) In the Apps Script editor, click Deploy, then Test deployments.
2) Create a new test deployment and select Gmail add-on.
3) Authorize when prompted.
4) Open Gmail and locate InboxCleaner in the right sidebar.

## Test in Gmail
- Open Gmail in the browser.
- Open the InboxCleaner add-on from the right sidebar.
- Use Preview to verify matches before cleaning.
- Use Clean now to trash or archive.

## Notes
- Messages are never permanently deleted. Trash or archive only.
- Only message IDs and metadata headers are used.
- The query always includes `-is:starred -is:important` exclusions.
