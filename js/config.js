/**
 * Swiss Club App — Configuration
 * ─────────────────────────────
 * Fill in the values below after completing Google Cloud setup.
 * This file is listed in .gitignore — do NOT commit real credentials.
 */

const CONFIG = {
  // Google OAuth 2.0 client ID (from Google Cloud Console → Credentials)
  CLIENT_ID: '778046723349-5mbg3pj5ed1vognpbi0i4b5cmm3ebc50.apps.googleusercontent.com',

  // Google Spreadsheet ID (from the sheet URL)
  SHEET_ID: '15C9IbCYjvkOW5USab3CraK0t1sX1LOYf',

  // Sheet tab names (must match exactly)
  SHEETS: {
    MEMBERS:      'Members',
    EVENTS:       'Events',
    TRANSACTIONS: 'Transactions',
  },

  // Whitelisted Google account emails — only these can log in
  ALLOWED_EMAILS: [
    'seandedios@gmail.com',
  ],

  // Google Apps Script RSVP relay URL (set after deploying rsvp-relay.gs)
  RSVP_RELAY_URL: '',

  // Base URL of your GitHub Pages site (no trailing slash)
  APP_BASE_URL: 'https://swisslinkph.github.io/swiss-link',

  // OAuth scopes
  SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
  ].join(' '),
};

// Freeze to prevent accidental mutation
Object.freeze(CONFIG);
Object.freeze(CONFIG.SHEETS);
