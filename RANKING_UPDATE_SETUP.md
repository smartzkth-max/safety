# Ranking Update Setup

This update ranks workers by:

1. Higher score first
2. If tied, shorter time taken wins
3. If still tied, earlier submitted time wins

## Files to upload to GitHub

Upload/replace:

- `src/App.jsx`
- `google-apps-script/Code.gs`
- `package.json`

Commit message:

```text
Add ranking by score and time taken
```

## Apps Script update

1. Open your Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Replace the code with `google-apps-script/Code.gs`.
4. Click **Save**.
5. Go to **Deploy → Manage deployments**.
6. Click the pencil/edit icon.
7. Select **New version**.
8. Click **Deploy**.

The existing `Responses` sheet will automatically add new columns:

- Started At
- Duration Seconds
- Duration

Old rows will not have duration, so exact tie-break ranking by time applies to new submissions after this update.
