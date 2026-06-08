# Swiss Club Member Portal

A private web app for Swiss Club administrators to manage members, events, front-desk check-ins, and email invitations. Hosted on GitHub Pages. Data stored in Google Sheets.

---

## Setup Guide

### 1. Google Cloud Project

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
2. Create a new project → name it **Swiss Club App**
3. Enable APIs:
   - **Google Sheets API**
   - **Gmail API**
4. Configure **OAuth consent screen**:
   - User type: **External**
   - App name: Swiss Club Portal
   - Add all admin email addresses as **Test users** (while the app is unverified)
   - Scopes to add:
     - `https://www.googleapis.com/auth/spreadsheets`
     - `https://www.googleapis.com/auth/gmail.send`
     - `https://www.googleapis.com/auth/userinfo.email`
     - `https://www.googleapis.com/auth/userinfo.profile`
5. Create **OAuth 2.0 Client ID**:
   - Application type: **Web application**
   - Authorised JavaScript origins:
     ```
     https://<your-github-username>.github.io
     ```
   - Authorised redirect URIs:
     ```
     https://<your-github-username>.github.io/swiss-club-app/
     https://<your-github-username>.github.io/swiss-club-app/index.html
     ```
6. Copy the **Client ID** — you'll need it shortly.

---

### 2. Google Apps Script RSVP Relay

1. Open your Google Sheet → **Extensions → Apps Script**
2. Create a new file, paste the contents of `scripts/rsvp-relay.gs`
3. Update `SHEET_ID` at the top to match your spreadsheet ID (already filled in if using the provided script)
4. **Deploy → New deployment**
   - Type: **Web App**
   - Execute as: **Me** (your Google account)
   - Who has access: **Anyone**
5. Click **Deploy**, then copy the **Web App URL**

---

### 3. Configure the App

Copy `js/config.js.template` to `js/config.js` and fill in your values:

```js
const CONFIG = {
  CLIENT_ID: 'YOUR_OAUTH_CLIENT_ID.apps.googleusercontent.com',
  API_KEY:    '',          // optional: only needed if making unauthenticated Sheets reads
  SHEET_ID:  '15C9IbCYjvkOW5USab3CraK0t1sX1LOYf',
  SHEETS: {
    MEMBERS:      'Members',
    EVENTS:       'Events',
    TRANSACTIONS: 'Transactions',
  },
  ALLOWED_EMAILS: [
    'admin@example.com',
    // add all authorised admin emails here
  ],
  RSVP_RELAY_URL: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
  APP_BASE_URL:   'https://<your-github-username>.github.io/swiss-club-app',
  SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
  ].join(' '),
};
```

> `js/config.js` is listed in `.gitignore` and will **not** be committed to GitHub.

---

### 4. GitHub Repository & Pages

1. Create a new GitHub repository named `swiss-club-app` (can be public — the app itself requires Google login)
2. Push all files:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/<username>/swiss-club-app.git
   git push -u origin main
   ```
   > **Note:** `config.js` is gitignored, so you must create it manually on any new machine.
3. Go to **Settings → Pages → Source**: `main` branch, `/ (root)` folder → Save
4. Your app will be live at: `https://<username>.github.io/swiss-club-app/`

---

### 5. Update Google OAuth Authorized Origins

Go back to your Google Cloud Console → **Credentials → your OAuth client** and confirm the authorised origins and redirect URIs match your live GitHub Pages URL exactly (including trailing slash).

---

## Usage

| Page | Purpose |
|---|---|
| `index.html` | Login with Google |
| `app.html#dashboard` | KPI overview, charts |
| `app.html#members` | Add / edit / delete / export members |
| `app.html#events` | Create and manage events |
| `app.html#frontdesk` | Door check-in and payment recording |
| `app.html#email` | Send event invite emails via Gmail |
| `rsvp.html` | Public RSVP page linked from invitation emails |

---

## Google Sheet Structure

| Tab | Columns |
|---|---|
| `Members` | Member Key, Full Name, First Name, Last Name, Email, Location, 2026 Membership Status, Membership Type, Family Group, … |
| `Events` | EventID, Title, Date, Location, Description, MemberFee, GuestFee, Status, CreatedDate, CreatedBy |
| `Transactions` | TransactionID, Timestamp, MemberKey, MemberName, EventID, EventName, AmountPaid, PaymentMode, Category, Year, HeadCount, Notes, RecordedBy |

The app will automatically create the **Events** and **Transactions** tabs with the correct headers on first login.

---

## Architecture Notes

- **No backend server** — all API calls go from the browser directly to Google's APIs
- **Auth** — Google OAuth 2.0 PKCE flow; token stored in `sessionStorage` (cleared on tab close)
- **Access control** — `ALLOWED_EMAILS` whitelist checked after OAuth login
- **RSVP** — lightweight token (`btoa(eventId:memberKey:sheetPrefix)`) verified server-side in the Apps Script relay
- **Emails** — sent from the logged-in admin's Gmail via the Gmail API; no third-party mail service needed
