# SheetsDark Chrome Extension

A simple, effective, and maintainable Chrome extension that enables dark mode on Google Sheets.

## What it does

- Targets only `https://docs.google.com/spreadsheets/*`
- Applies a dark theme using a lightweight CSS strategy
- Includes an extension popup to enable/disable dark mode
- Persists preference using `chrome.storage.sync`

## Project structure

- `manifest.json` - Extension configuration (Manifest V3)
- `src/content.js` - Reads state and toggles dark mode attribute
- `src/dark-theme.css` - Theme rules gated by `data-sheets-dark`
- `popup.html` - Small popup UI
- `popup.js` - Popup logic and persisted toggle

## Load in Chrome

1. Open Chrome and go to `chrome://extensions`.
2. Turn on **Developer mode**.
3. Click **Load unpacked**.
4. Select this folder: `SheetsDark`.

## Usage

1. Open any Google Sheets document.
2. Click the extension icon.
3. Toggle **Enable dark mode** on or off.
4. Refresh the tab if the current document does not update immediately.

## Notes

- This extension intentionally keeps logic small and centralized for easy updates.
- If Google changes Sheets DOM behavior, tune `src/dark-theme.css` selectors without changing extension architecture.
