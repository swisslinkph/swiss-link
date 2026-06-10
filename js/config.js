/**
 * Swiss Club App — Configuration
 * ─────────────────────────────
 * Fill in the values below after completing Google Cloud setup.
 * This file is listed in .gitignore — do NOT commit real credentials.
 */

const CONFIG = {
  // Google OAuth 2.0 client ID (from Google Cloud Console → Credentials)
  CLIENT_ID: '778046723349-5mbg3pj5ed1vognpbi0i4b5cmm3ebc50.apps.googleusercontent.com',
  CLIENT_SECRET: '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',

  // Google Spreadsheet ID (from the sheet URL)
  SHEET_ID: '1xOs_Gii-7N-y97wdaR5sr-5ancGQqz7BA2YNF85UauA',

  // Sheet tab names (must match exactly)
  SHEETS: {
    MEMBERS:      'Members',
    EVENTS:       'Events',
    TRANSACTIONS: 'Transactions',
    ADMINS:       'Admins',
  },

  // Whitelisted Google account emails — only these can log in
  ALLOWED_EMAILS: [
    'seandedios@gmail.com',
    'swisslinkph@gmail.com',
    'jleperalta@gmail.com',
    'jonas.oberle@gmail.com',
    'klegifam@gmail.com',
    'sevket.oezdes@gmail.com',
    'swissyesyes@gmail.com',
  ],

  // Google Apps Script RSVP relay URL (set after deploying rsvp-relay.gs)
  RSVP_RELAY_URL: 'https://script.google.com/macros/s/AKfycbzUBdL9ASOANUEAib_3XXXlp97y381pyKh4Ap2xyndmaxRJGTeOwjg9hBLVNnaKaBXw/exec',

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
