# Google Sheets Backend Setup

This app uses a Google Sheet as a simple results database.

## 1. Create the Google Sheet

1. Open Google Drive.
2. Create a new Google Sheet.
3. Name it: `Safety Quiz Results`.
4. You do not need to create columns manually. The script creates a `Responses` sheet and headers automatically.

## 2. Add Apps Script

1. In the Google Sheet, click **Extensions**.
2. Click **Apps Script**.
3. Delete any starter code.
4. Paste the code from `google-apps-script/Code.gs`.
5. Click **Save**.

## 3. Deploy as Web App

1. In Apps Script, click **Deploy**.
2. Click **New deployment**.
3. Click the gear icon / Select type.
4. Choose **Web app**.
5. Set:
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy**.
7. Authorize the script.
8. Copy the Web App URL ending with `/exec`.

## 4. Use in the safety app

1. Open organizer page:
   `https://smartzkth-max.github.io/safety/`
2. Paste the `/exec` URL into **Google Apps Script Web App URL**.
3. The QR code will include the API URL.
4. Workers scan the QR and submit answers.
5. Organizer enters PIN `1234`, then clicks **Refresh results**.

## Notes

- Worker mode URL contains `&mode=worker`, so workers do not see QR generator or admin panel.
- This is a practical prototype. For stronger security, replace the demo PIN and add a server-side organizer token.
