# MessMate

MessMate is a mobile-first, local-first mess manager built with plain HTML, CSS, and JavaScript.

## Run

Open `index.html` in a current browser. Secure local encryption uses the Web Crypto API, which requires a secure context; use HTTPS or `localhost` if opening the file directly is not supported by the browser.

## Data and settings

- Each mess is encrypted separately with its own password and stored in that browser's local storage.
- Use Settings to export an encrypted JSON backup for transfer to another device.
- App name, icon, favicon, and text size preferences are stored locally in the current browser.
- No server, build step, or third-party runtime dependency is required.
