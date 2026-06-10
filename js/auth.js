/**
 * auth.js — Google OAuth 2.0 with PKCE (client-side, no server needed)
 * Handles sign-in, sign-out, token storage, and email whitelist check.
 */

const Auth = (() => {
  // ── Storage keys ─────────────────────────────────────────────────────────
  const KEY_TOKEN       = 'sc_access_token';
  const KEY_EXPIRY      = 'sc_token_expiry';
  const KEY_USER_EMAIL  = 'sc_user_email';
  const KEY_USER_NAME   = 'sc_user_name';
  const KEY_USER_PIC    = 'sc_user_pic';
  const KEY_CODE_VERIF  = 'sc_code_verifier';

  let _dynamicAdmins = []; // emails loaded from Admins sheet at login

  // ── Internal helpers ─────────────────────────────────────────────────────
  function b64url(buf) {
    return btoa(String.fromCharCode(...new Uint8Array(buf)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  }

  async function generatePKCE() {
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const encoder  = new TextEncoder();
    const digest   = await crypto.subtle.digest('SHA-256', encoder.encode(verifier));
    const challenge = b64url(digest);
    return { verifier, challenge };
  }

  async function fetchUserInfo(token) {
    const res  = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${token}` },
    });
    return res.json();
  }

  function isAllowed(email) {
    const e = email.toLowerCase();
    if (!CONFIG.ALLOWED_EMAILS.length && !_dynamicAdmins.length) return true;
    return CONFIG.ALLOWED_EMAILS.includes(e) || _dynamicAdmins.includes(e);
  }

  async function _fetchAdminEmails(token) {
    try {
      const sheet = (CONFIG.SHEETS && CONFIG.SHEETS.ADMINS) || 'Admins';
      const range = encodeURIComponent(`${sheet}!A:A`);
      const url   = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}`;
      const res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) return [];
      const data  = await res.json();
      return (data.values || []).slice(1)
        .map(r => (r[0] || '').toLowerCase().trim())
        .filter(Boolean);
    } catch {
      return [];
    }
  }

  // ── Public API ────────────────────────────────────────────────────────────
  async function signIn() {
    const { verifier, challenge } = await generatePKCE();
    sessionStorage.setItem(KEY_CODE_VERIF, verifier);

    const params = new URLSearchParams({
      client_id:             CONFIG.CLIENT_ID,
      redirect_uri:          window.location.origin + window.location.pathname,
      response_type:         'code',
      scope:                 CONFIG.SCOPES,
      code_challenge:        challenge,
      code_challenge_method: 'S256',
      access_type:           'offline',
      prompt:                'select_account',
    });

    window.location.href = `https://accounts.google.com/o/oauth2/v2/auth?${params}`;
  }

  async function handleCallback() {
    const url    = new URL(window.location.href);
    const code   = url.searchParams.get('code');
    const error  = url.searchParams.get('error');

    if (error) throw new Error(`OAuth error: ${error}`);
    if (!code)  return false; // not a callback

    const verifier = sessionStorage.getItem(KEY_CODE_VERIF);
    if (!verifier) throw new Error('Missing PKCE verifier');

    // Exchange code for token
    const body = new URLSearchParams({
      client_id:     CONFIG.CLIENT_ID,
      client_secret: CONFIG.CLIENT_SECRET,
      code,
      code_verifier: verifier,
      grant_type:    'authorization_code',
      redirect_uri:  window.location.origin + window.location.pathname,
    });

    const res  = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
    });

    const data = await res.json();
    if (data.error) throw new Error(data.error_description || data.error);

    // Get user profile
    const profile = await fetchUserInfo(data.access_token);

    _dynamicAdmins = await _fetchAdminEmails(data.access_token);
    if (!isAllowed(profile.email)) {
      throw new Error(`Access denied: ${profile.email} is not an authorised admin.`);
    }

    // Persist to sessionStorage
    const expiry = Date.now() + (data.expires_in - 60) * 1000;
    sessionStorage.setItem(KEY_TOKEN,      data.access_token);
    sessionStorage.setItem(KEY_EXPIRY,     expiry);
    sessionStorage.setItem(KEY_USER_EMAIL, profile.email);
    sessionStorage.setItem(KEY_USER_NAME,  profile.name || profile.email);
    sessionStorage.setItem(KEY_USER_PIC,   profile.picture || '');

    // Clean URL
    const clean = window.location.pathname + window.location.hash;
    window.history.replaceState({}, '', clean);

    return true;
  }

  function isAuthenticated() {
    const token  = sessionStorage.getItem(KEY_TOKEN);
    const expiry = parseInt(sessionStorage.getItem(KEY_EXPIRY) || '0', 10);
    return !!token && Date.now() < expiry;
  }

  function getToken()     { return sessionStorage.getItem(KEY_TOKEN); }
  function getUserEmail() { return sessionStorage.getItem(KEY_USER_EMAIL) || ''; }
  function getUserName()  { return sessionStorage.getItem(KEY_USER_NAME)  || ''; }
  function getUserPic()   { return sessionStorage.getItem(KEY_USER_PIC)   || ''; }

  function signOut() {
    [KEY_TOKEN, KEY_EXPIRY, KEY_USER_EMAIL, KEY_USER_NAME,
     KEY_USER_PIC, KEY_CODE_VERIF].forEach(k => sessionStorage.removeItem(k));
    window.location.href = 'index.html';
  }

  function setAdminEmails(emails) {
    _dynamicAdmins = emails.map(e => (e || '').toLowerCase().trim()).filter(Boolean);
  }

  return { signIn, handleCallback, isAuthenticated, getToken,
           getUserEmail, getUserName, getUserPic, signOut, setAdminEmails };
})();
