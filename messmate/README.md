# MessMate

MessMate is a mobile-first, local-first mess manager built with plain HTML, CSS, and JavaScript.

## Run

Open `index.html` in a current browser for local-only use. Secure local encryption uses the Web Crypto API, which requires a secure context; use HTTPS or `localhost` if opening the file directly is not supported by the browser. Google Sheets sign-in specifically requires HTTPS or localhost; a `file://` page cannot use OAuth.

## Data and settings

- Each mess is encrypted separately with its own password and stored in that browser's local storage.
- Use Settings to export an encrypted JSON backup for transfer to another device.
- App name, icon, favicon, and text size preferences are stored locally in the current browser.
- Local-only use requires no server, build step, or third-party runtime dependency. Optional Google Sheets sync uses Google's Identity Services library and Sheets API.

## Optional Google Sheets sync

Google Sheets holds an encrypted backup of the active mess. Local storage remains the source of truth; each successful local save uploads the encrypted snapshot, and the status indicator turns green only after Google confirms it. Offline or failed uploads stay local and can be retried with **Sync Now** or by reconnecting. The app reads the sheet only when you explicitly choose **Restore from Google Sheets**. Restore copies a selected snapshot to local storage and never overwrites existing local mess data. Devices do not merge: a browser that has not synced a mess before appends its own copy rather than finding or overwriting another device's row.

To connect:

1. In Google Cloud Console, create a project, enable the Google Sheets API, and create an OAuth Client ID for a Web application.
2. Add the deployed app's HTTPS origin to the OAuth client's Authorized JavaScript origins. Use `http://localhost:<port>` for local development; OAuth cannot use a `file://` origin.
3. Create a Google spreadsheet and a tab named `MessMate`. Put this exact header row in row 1: `Mess ID | Chunk | Chunks | Mess Name | Updated At | Salt | IV | Encrypted Payload`.
4. Open MessMate Settings, paste the OAuth Client ID and spreadsheet URL or ID, save, then connect a Google account that has edit access to the spreadsheet. If the OAuth app is in testing mode, add that account as a test user.
5. To recover after losing a device, open MessMate on the new device, choose **Restore from Google Sheets**, enter the same OAuth Client ID and spreadsheet URL or ID, connect to Google, select the newest snapshot, then unlock it with the original mess password. Once it is unlocked, later local changes sync as usual. Alternatively, use the encrypted JSON backup.

Google stores encrypted snapshots in chunks, so larger messes do not depend on a single-cell payload. The mess password is not sent to Google. OAuth access is granted to the signed-in Google account and token refresh may require reconnecting. A green indicator confirms the latest acknowledged upload; if configuration or sign-in is missing, the app continues to work locally and shows a non-green status.

The OAuth Client ID and spreadsheet ID are identifiers, not passwords, and a browser app's Client ID is necessarily visible. Do not add a Google OAuth client secret to this static app. Anyone still needs Google permission to access the spreadsheet; keep it private and grant access only to trusted accounts. This app requests the Sheets OAuth scope, so use a dedicated Google account for the mess if you want to limit which of your account's spreadsheets the granted token could access.
