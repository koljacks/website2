/**
 * ═══════════════════════════════════════════════════════════════════
 *  CMS ENTERPRISE PORTAL v3.6
 *  Callback Management System — Role-based WFM tooling
 *  Built by: Jackson Koli (WFM RTA, RGS Mumbai)
 * ═══════════════════════════════════════════════════════════════════
 *
 *  SETUP (once):
 *    1. File → Project Settings → Timezone: (GMT+05:30) India
 *    2. Run setupCMSv3() from editor (authorize)
 *    3. Run installAllTriggers()
 *    4. Deploy → New deployment → Web app
 *       ▸ Execute as: Me
 *       ▸ Who has access: Anyone within <your Google Workspace domain>
 *
 *  v3.6 CHANGES FROM v3.5:
 *    ▸ Email relay through Account B (UrlFetchApp → relay web app)
 *    ▸ 6 structured email types only (Reminder, L1, L2, Complete, Reschedule, Reset)
 *    ▸ Reminder: TL in To, Manager in CC
 *    ▸ L1: Both TL and Manager in To
 *    ▸ L2: Manager in To, TL in CC
 *    ▸ HTML emails with branded template (buildEmailHtml_)
 *    ▸ Escalation tracker via Script Properties (no duplicate L1/L2 sends)
 *    ▸ Week runs Sunday → Saturday in getAgentStats
 *    ▸ All duplicate functions removed, all bugs fixed
 * ═══════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════
//  SECTION 0 — EMAIL RELAY CONFIG (Account B)
// ═══════════════════════════════════════════════════════════════════

var EMAIL_RELAY_URL    = 'https://script.google.com/macros/s/AKfycbyeO4Y3SERwWSfN5d_nPlygNzp5b7dw1dR-jXIkAUE7NEcI-IMoEDjQH4nMpC2-NDYm1Q/exec';
var EMAIL_RELAY_SECRET = 'cms_tp_relay_2024'; // must match relay script on Account B


// ═══════════════════════════════════════════════════════════════════
//  SECTION 1 — CONFIG
// ═══════════════════════════════════════════════════════════════════

const CFG_CACHE_KEY = 'cms_config_v3';
const HC_CACHE_KEY  = 'cms_headcount_v3';
const CFG_TTL = 300;
const HC_TTL  = 21600;


function defaultConfig_() {
  return {
    TIMEZONE: 'Asia/Kolkata',
    SECONDARY_SHEET_ID: '',
    CMS_SHEET: 'CMS',
    HEADCOUNT_SHEET: 'Headcount',
    AUDIT_SHEET: 'AuditLog',
    ERROR_SHEET: 'ErrorLog',
    REMINDER_WINDOW_MIN: 5,
    REMINDER_TOLERANCE_MIN: 1.5,
    REMINDER_EARLY_MIN: 15,
    OVERDUE_THRESHOLD_MIN: 10,
    MAX_DAILY_ENTRIES_PER_AGENT: 200,
    DUPLICATE_WINDOW_HOURS: 24,
    HIGH_PRIORITY_REASONS: 'Mediation,On-trip Issues,Bug update',
    EMAIL_ON_PRIORITY: 'High',
    DASHBOARD_DAYS_DEFAULT: 30,
    DASHBOARD_DAYS_MAX: 180,
    NOTIFICATION_POLL_SECONDS: 30,
    CIRCUIT_BREAKER_THRESHOLD: 3,
    CIRCUIT_BREAKER_PAUSE_MIN: 60,
    APP_VERSION: 'v3.6',
    DAILY_TARGET_PER_AGENT: 8,
    PORTAL_URL: 'YOUR_PORTAL_URL_HERE'
  };
}

const MAX_FAILED_ATTEMPTS = 5;
const LOCKOUT_MINUTES = 15;
const DASHBOARD_QUERY_TIMEOUT_MS = 15000;
const DASHBOARD_MAX_ROWS = 2000;

// ── SESSION TOKEN TTL (8 hours — matches frontend) ──────────────────
const SESSION_TOKEN_TTL_MS = 8 * 3600 * 1000;

function getConfig() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CFG_CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const cfg = defaultConfig_();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim();
      if (k) cfg[k] = data[i][1];
    }
  }
  cache.put(CFG_CACHE_KEY, JSON.stringify(cfg), CFG_TTL);
  return cfg;
}

function refreshConfigCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CFG_CACHE_KEY);
  cache.remove(HC_CACHE_KEY); // ADD THIS
  getConfig();
  SpreadsheetApp.getActive().toast('Caches cleared.', 'CMS v3.6', 3);
}

// ═══════════════════════════════════════════════════════════════════
//  SECTION 2 — HEADCOUNT + ROLE DERIVATION
// ═══════════════════════════════════════════════════════════════════

/**
 * Headcount col layout:
 *   A=EmpCode  B=Name  C=Supervisor  D=Manager  E=SkipManager
 *   F=Designation  G=LOB  H=AdvisorEmail  I=SupervisorEmail  J=ManagerEmail
 *   K=Role (optional override)
 */
const HC_CHUNK_PREFIX = 'cms_hc_chunk_v3_';
const HC_META_KEY     = 'cms_hc_meta_v3';
const HC_CHUNK_SIZE   = 400;

function getHeadcountMap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(HC_CACHE_KEY);
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }

  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.HEADCOUNT_SHEET);
  if (!sheet) throw new Error('Headcount sheet not found: ' + cfg.HEADCOUNT_SHEET);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const emp = String(data[i][0] || '').trim();
    if (!emp) continue;
    map[emp] = {
      emp, name: data[i][1] || '',
      supervisor: String(data[i][2] || '').trim(),
      manager:    String(data[i][3] || '').trim(),
      skipManager: data[i][4] || '',
      designation: data[i][5] || '',
      lob:         data[i][6] || '',
      advisorEmail:    String(data[i][7] || '').trim().toLowerCase(),
      supervisorEmail: String(data[i][8] || '').trim().toLowerCase(),
      managerEmail:    String(data[i][9] || '').trim().toLowerCase(),
      email: String(data[i][7] || '').trim().toLowerCase(),
      role: deriveRole_(data[i][5], data[i][10])
    };
  }
  try { cache.put(HC_CACHE_KEY, JSON.stringify(map), HC_TTL); } catch(e) {}
  return map;
}
function deriveRole_(designation, explicit) {
  const override = String(explicit || '').trim();
  if (override) return override;
  const d = String(designation || '').toLowerCase().trim();
  if (/\bwfm\b|\brta\b/.test(d)) return 'WFM';
  if (/\b(quality\s*analyst|qa\s*analyst|quality\s*assurance)\b/.test(d)) return 'QA';
  if (/(operations?|support)/.test(d) && /\b(manager|director|vice\s*president|vp|president)\b/.test(d)) return 'TM';
  if (/(team leader|deputy team leader)/.test(d) &&
      (/(operations?|support)/.test(d) || /^(deputy )?team leader$/i.test(d))) return 'TL';
  return 'Agent';
}

/**
 * Given an agent emp code, returns their TL and Manager profile objects.
 * Looks up by supervisor email (column I) and manager email (column J).
 */
/**
 * Returns TL and Manager objects for an agent.
 * Headcount col I = Team Leader email (supervisorEmail)
 * Headcount col J = Manager email (managerEmail)
 *
 * Tries to find the full profile from the headcount map by email match.
 * Falls back to a minimal object with just name + email if not found
 * (handles cases where TL/Manager are not in the headcount sheet).
 */
function getTlAndManager_(agentEmp) {
  const hcMap  = getHeadcountMap_();
  const agent  = hcMap[String(agentEmp || '').trim()];
  if (!agent) return { tl: null, manager: null };

  // TL email is already in col I (supervisorEmail) — use it directly
  const tlEmail  = String(agent.supervisorEmail  || '').trim().toLowerCase();
  const mgrEmail = String(agent.managerEmail     || '').trim().toLowerCase();

  // Try to find full profile from headcount map (gives us name, role, etc.)
  let tl = tlEmail
    ? (Object.values(hcMap).find(function(p) { return p.advisorEmail === tlEmail; }) ||
       { advisorEmail: tlEmail, name: agent.supervisor || 'Team Leader' })
    : null;

  let manager = mgrEmail
    ? (Object.values(hcMap).find(function(p) { return p.advisorEmail === mgrEmail; }) ||
       { advisorEmail: mgrEmail, name: agent.manager || 'Manager' })
    : null;

  return { tl: tl, manager: manager };
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 3 — WEB APP ENTRY + AUTH
// ═══════════════════════════════════════════════════════════════════

function doGet() {
  const tmpl = HtmlService.createTemplateFromFile('WebPage');
  tmpl.version = getConfig().APP_VERSION;
  return tmpl.evaluate()
    .setTitle('CMS Enterprise Portal ' + getConfig().APP_VERSION)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function login(empId) {
  empId = String(empId || '').trim();
  if (!empId) throw new Error('Employee ID is required.');
  
  const profile = getHeadcountMap_()[empId];
  if (!profile) throw new Error('Employee ID not found: ' + empId);

  if (profile.role === 'Agent') {
    audit_('LOGIN', empId, '', 'role=Agent');
    return {
      emp: profile.emp,
      name: profile.name,
      designation: profile.designation,
      lob: profile.lob,
      advisorEmail: profile.advisorEmail,
      supervisor: profile.supervisor,
      manager: profile.manager,
      supervisorEmail: profile.supervisorEmail,
      managerEmail: profile.managerEmail,
      role: profile.role,
      serverTime: new Date().toISOString(),
      upcoming: getUpcomingForAgent_(empId, 10),
      todayCount: countEntriesToday_(empId),
      dailyTarget: Number(getConfig().DAILY_TARGET_PER_AGENT) || 8,
      pendingReassigns: getPendingReassignsForAgent_(empId),
      sessionToken: generateSessionToken_(empId)
    };
  }

  const cred = getAdminCredential_(empId);
  const locked = cred && isLocked_(cred);
  return {
    needsPassword: true,
    emp: profile.emp,
    name: profile.name,
    designation: profile.designation,
    role: profile.role,
    firstTime: !cred,
    locked: !!locked,
    lockedUntil: locked ? cred.lockedUntil : ''
  };
}

function completeAdminLogin(empId, password, newPassword) {
  // Validate inputs
  if (!empId) throw new Error('Employee ID is required.');
  
  empId = String(empId).trim();
  const profile = getHeadcountMap_()[empId];
  if (!profile) throw new Error('Employee ID not found.');
  if (profile.role === 'Agent') throw new Error('Agents log in with emp code only.');

  const cred = getAdminCredential_(empId);

  // First time login
  if (!cred) {
    if (!newPassword) throw new Error('First-time login: please set a password (min 8 chars, letters + numbers).');
    validatePassword_(newPassword);
    setAdminCredential_(empId, newPassword);
    audit_('PASSWORD_SET_FIRSTTIME', empId, '', 'role=' + profile.role);
    
    // Return minimal response - dashboard will load async
    return {
      emp: profile.emp,
      name: profile.name,
      designation: profile.designation || '',
      lob: profile.lob || '',
      advisorEmail: profile.advisorEmail || '',
      supervisor: profile.supervisor || '',
      manager: profile.manager || '',
      supervisorEmail: profile.supervisorEmail || '',
      managerEmail: profile.managerEmail || '',
      role: profile.role,
      serverTime: new Date().toISOString(),
      // Empty dashboard - will be loaded separately
      dashboard: { stats: { pending: 0, completed: 0, overdue: 0, today: 0, noAnswer: 0 }, entries: [], agents: [], lobs: [] },
      config: [],
      todayCount: 0,
      dailyTarget: 8,
      upcoming: [],
      sessionToken: generateSessionToken_(empId)
    };
  }

  // Check locked account
  if (isLocked_(cred)) {
    audit_('LOGIN_BLOCKED_LOCKED', empId, '', 'until=' + cred.lockedUntil);
    throw new Error('Account locked until ' + cred.lockedUntil + '. Contact WFM to unlock.');
  }

  // Verify password
  if (!verifyPassword_(password, cred)) {
    incrementFailedAttempts_(empId, cred);
    const newCount = (cred.failedAttempts || 0) + 1;
    const remaining = MAX_FAILED_ATTEMPTS - newCount;
    audit_('LOGIN_FAILED', empId, '', 'attempt=' + newCount);
    if (remaining <= 0) throw new Error('Too many failed attempts. Account locked for ' + LOCKOUT_MINUTES + ' minutes.');
    throw new Error('Incorrect password. ' + remaining + ' attempt' + (remaining === 1 ? '' : 's') + ' remaining.');
  }

  // Successful login
  resetFailedAttempts_(empId);
  audit_('LOGIN', empId, '', 'role=' + profile.role);
  
  // Return minimal response - dashboard will load async
  return {
    emp: profile.emp,
    name: profile.name,
    designation: profile.designation || '',
    lob: profile.lob || '',
    advisorEmail: profile.advisorEmail || '',
    supervisor: profile.supervisor || '',
    manager: profile.manager || '',
    supervisorEmail: profile.supervisorEmail || '',
    managerEmail: profile.managerEmail || '',
    role: profile.role,
    serverTime: new Date().toISOString(),
    // Empty dashboard - will be loaded separately
    dashboard: { stats: { pending: 0, completed: 0, overdue: 0, today: 0, noAnswer: 0 }, entries: [], agents: [], lobs: [] },
    config: [],
    todayCount: 0,
    dailyTarget: 8,
    upcoming: [],
    sessionToken: generateSessionToken_(empId)
  };
}
function loadAdminDashboard(actorEmp, daysFilter) {
  console.log('loadAdminDashboard called for:', actorEmp);
  try {
    const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
    const dashboard = buildScopedDashboard_(actor, daysFilter || 30);
    console.log('Dashboard built successfully, pending:', dashboard.stats.pending);
    return { success: true, dashboard: dashboard };
  } catch(e) {
    console.error('loadAdminDashboard error:', e.message);
    return { success: false, error: e.message };
  }
}
function testHeadcountSheet() {
  try {
    const cfg = getConfig();
    const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.HEADCOUNT_SHEET);
    if (!sheet) {
      console.log('Headcount sheet not found');
      return;
    }
    const lastRow = sheet.getLastRow();
    console.log('Headcount sheet rows:', lastRow);
    
    // Test reading a specific employee
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < Math.min(20, data.length); i++) {
      console.log('Row ' + i + ': Emp=' + data[i][0] + ', Name=' + data[i][1]);
    }
    
    return { success: true, rowCount: lastRow };
  } catch(e) {
    console.error('testHeadcountSheet error:', e.message);
    return { error: e.message };
  }
}
function buildAdminLoginPayload_(profile) {
  var dash;
  try {
    dash = buildScopedDashboard_(profile, 30);
  } catch (e) {
    dash = { stats: { pending: 0, completed: 0, overdue: 0, today: 0, noAnswer: 0 }, entries: [], agents: [], lobs: [], daysFilter: 30, queryMs: 0, partial: false, repeatOffenders: [] };
  }
  return {
    // Include ALL profile fields that the frontend expects
    emp: profile.emp,
    name: profile.name,
    designation: profile.designation,
    lob: profile.lob,
    advisorEmail: profile.advisorEmail,
    supervisor: profile.supervisor,
    manager: profile.manager,
    supervisorEmail: profile.supervisorEmail,
    managerEmail: profile.managerEmail,
    role: profile.role,
    viewMode: profile.role === 'WFM' ? 'WFM_ADMIN' : 'STANDARD_ADMIN',
    dashboard: dash,
    config: getConfigForAdmin_(),
    serverTime: new Date().toISOString()
    // roster intentionally omitted — loaded async by loadRoster() in frontend
  };
}

// ── SESSION TOKENS ────────────────────────────────────────────────────
// Lets agents and admins survive a page refresh without re-logging in.
// Tokens are UUID strings stored in Script Properties keyed by emp ID.
// TTL = 8 hours.  On logout, the client just clears localStorage —
// the server token expires naturally on its own after 8h.

function generateSessionToken_(empId) {
  var token    = Utilities.getUuid().replace(/-/g, '');
  var props    = PropertiesService.getScriptProperties();
  var sessions = {};
  try { sessions = JSON.parse(props.getProperty('cms_sessions') || '{}'); } catch(e) {}
  // Prune expired entries while we have the object open
  var now = Date.now();
  Object.keys(sessions).forEach(function(k) {
    if (now - sessions[k].ts > SESSION_TOKEN_TTL_MS) delete sessions[k];
  });
  sessions[String(empId)] = { token: token, ts: now };
  props.setProperty('cms_sessions', JSON.stringify(sessions));
  return token;
}

function verifySessionToken_(empId, token) {
  if (!empId || !token) return false;
  try {
    var s = JSON.parse(PropertiesService.getScriptProperties().getProperty('cms_sessions') || '{}');
    var e = s[String(empId)];
    return !!(e && e.token === String(token) && (Date.now() - e.ts) <= SESSION_TOKEN_TTL_MS);
  } catch(e) { return false; }
}

/**
 * Called by the frontend on page refresh for non-Agent roles.
 * Verifies the session token stored in the browser's localStorage.
 * If valid → returns the minimal profile + a freshly rolled token (sliding 8h window).
 * If invalid/expired → throws 'SESSION_EXPIRED' so the frontend shows the login form.
 */
function restoreAdminSession(empId, token) {
  empId = String(empId || '').trim();
  if (!empId || !token)                  throw new Error('SESSION_EXPIRED');
  if (!verifySessionToken_(empId, token)) throw new Error('SESSION_EXPIRED');
  const profile = getHeadcountMap_()[empId];
  if (!profile) throw new Error('Employee not found.');
  if (profile.role === 'Agent') throw new Error('Agents use login() for session restore.');
  audit_('SESSION_RESTORE', empId, '', 'role=' + profile.role);
  var newToken = generateSessionToken_(empId); // rolling refresh
  return {
    emp: profile.emp, name: profile.name,
    designation: profile.designation || '', lob: profile.lob || '',
    advisorEmail: profile.advisorEmail || '', supervisor: profile.supervisor || '',
    manager: profile.manager || '', supervisorEmail: profile.supervisorEmail || '',
    managerEmail: profile.managerEmail || '', role: profile.role,
    serverTime: new Date().toISOString(),
    dashboard: { stats: { pending:0, completed:0, overdue:0, today:0, noAnswer:0 }, entries:[], agents:[], lobs:[] },
    config: [], todayCount: 0, dailyTarget: 8, upcoming: [],
    sessionToken: newToken
  };
}

function changeMyPassword(empId, currentPassword, newPassword) {
  const profile = assertRole_(empId, ['TL', 'TM', 'WFM']);
  const cred    = getAdminCredential_(empId);
  if (!cred) throw new Error('No existing credential found.');
  if (!verifyPassword_(currentPassword, cred)) {
    audit_('PASSWORD_CHANGE_FAILED', empId, '', 'bad current pw');
    throw new Error('Current password is incorrect.');
  }
  validatePassword_(newPassword);
  setAdminCredential_(empId, newPassword);
  audit_('PASSWORD_CHANGED', empId, '', 'self');
  return { ok: true };
}

function adminResetPassword(actorEmp, targetEmp) {
  const actor  = assertRole_(actorEmp, ['WFM']);
  const target = getHeadcountMap_()[String(targetEmp || '').trim()];
  if (!target) throw new Error('Target employee not found.');
  if (target.role === 'Agent') throw new Error('Agents do not have passwords.');
  deleteAdminCredential_(targetEmp);
  audit_('PASSWORD_RESET', actorEmp, '', 'target=' + targetEmp + ' name=' + target.name);

  // Send password reset email
  if (target.advisorEmail) {
    sendEmailSafe_({
      to:      target.advisorEmail,
      cc:      '',
      subject: '\uD83D\uDD11 Your CMS password has been reset',
      html: buildEmailHtml_({
        headerColor:    '#2C2C2A',
        title:          '\uD83D\uDD11 Password Reset',
        tagline:        'A WFM Admin has reset your CMS portal password.',
        greeting:       'Hi ' + target.name.split(' ')[0] + ',',
        alertMsg:       'Your CMS portal password has been reset. <b>Please log in and set a new password immediately.</b> Your account is locked until you do so.',
        alertColor:     '#EEEDFE', alertBorder: '#534AB7', alertTextColor: '#3C3489',
        rows: [
          ['Account',         target.advisorEmail],
          ['Reset By',        actor.name + ' (WFM Admin)'],
          ['Reset At',        fmt_(new Date(), 'dd-MM-yyyy HH:mm') + ' IST'],
          ['Action Required', 'Log in and set a new password'],
        ],
        cta:  'Log In & Set New Password',
        note: 'If you did not request this reset, contact your WFM Admin immediately.',
      })
    });
  }

  return { ok: true, message: target.name + ' will set a new password on next login.' };
}

function adminUnlockAccount(actorEmp, targetEmp) {
  assertRole_(actorEmp, ['WFM']);
  const cred = getAdminCredential_(targetEmp);
  if (!cred) throw new Error('No credential exists for that employee.');
  const sheet = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  sheet.getRange(cred.rowNum, 6, 1, 2).setValues([[0, '']]);
  audit_('ACCOUNT_UNLOCKED', actorEmp, '', 'target=' + targetEmp);
  return { ok: true };
}

function adminListCredentials(actorEmp) {
  assertRole_(actorEmp, ['WFM']);
  const sheet = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  const hcMap = getHeadcountMap_();
  const rows  = [];
  const seen  = {};
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const emp = String(data[i][0] || '').trim();
      if (!emp) continue;
      seen[emp] = true;
      const hc          = hcMap[emp] || {};
      const lockedUntil = data[i][6] || '';
      rows.push({
        emp: emp, name: hc.name || '(unknown)', role: hc.role || '?', lob: hc.lob || '',
        created: data[i][3] || '', lastLogin: data[i][4] || '',
        failedAttempts: Number(data[i][5] || 0), lockedUntil: lockedUntil,
        locked: !!lockedUntil && new Date(parseDdMmYyyy_(lockedUntil)).getTime() > Date.now(),
        passwordSet: true
      });
    }
  }
  Object.values(hcMap).forEach(p => {
    if (p.role === 'Agent' || seen[p.emp]) return;
    rows.push({ emp: p.emp, name: p.name, role: p.role, lob: p.lob,
      created: '', lastLogin: '', failedAttempts: 0, lockedUntil: '', locked: false, passwordSet: false });
  });
  return rows.sort((a, b) => (a.role + a.name).localeCompare(b.role + b.name));
}

function assertRole_(empId, allowed) {
  const profile = getHeadcountMap_()[String(empId || '').trim()];
  if (!profile) throw new Error('Employee not found.');
  if (allowed.indexOf(profile.role) === -1) {
    throw new Error('This action requires ' + allowed.join('/') + ' role. You are ' + profile.role + '.');
  }
  return profile;
}


// ═══════════════════════════════════════════════════════════════════
//  PASSWORD HELPERS
// ═══════════════════════════════════════════════════════════════════

function validatePassword_(password) {
  if (!password || typeof password !== 'string') throw new Error('Password required.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');
  if (password.length > 100) throw new Error('Password too long (max 100).');
  if (!/[a-zA-Z]/.test(password) || !/[0-9]/.test(password)) {
    throw new Error('Password must contain at least one letter and one number.');
  }
}

function generateSalt_() {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + '-' + Date.now() + '-' + Math.random())
  ).substring(0, 24);
}

function hashPassword_(password, salt) {
  let h = String(password) + '|' + String(salt);
  for (let i = 0; i < 5000; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h));
  }
  return h;
}

function verifyPassword_(password, cred) {
  if (!cred || !cred.hash || !cred.salt) return false;
  return hashPassword_(password, cred.salt) === cred.hash;
}

function getAdminCredential_(empId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data   = sheet.getDataRange().getValues();
  const needle = String(empId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === needle) {
      return { rowNum: i + 1, empId: needle,
        hash: String(data[i][1] || ''), salt: String(data[i][2] || ''),
        created: data[i][3] || '', lastLogin: data[i][4] || '',
        failedAttempts: Number(data[i][5] || 0), lockedUntil: String(data[i][6] || '') };
    }
  }
  return null;
}

function setAdminCredential_(empId, password) {
  const salt     = generateSalt_();
  const hash     = hashPassword_(password, salt);
  const now      = fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss');
  const sheet    = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  if (!sheet) throw new Error('AdminCredentials sheet missing. Run setupCMSv3().');
  const existing = getAdminCredential_(empId);
  if (existing) {
    sheet.getRange(existing.rowNum, 1, 1, 7).setValues([[empId, hash, salt, existing.created || now, now, 0, '']]);
  } else {
    sheet.appendRow([empId, hash, salt, now, now, 0, '']);
  }
}

function deleteAdminCredential_(empId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  const cred  = getAdminCredential_(empId);
  if (cred) sheet.deleteRow(cred.rowNum);
}

function incrementFailedAttempts_(empId, cred) {
  const sheet    = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  const newCount = (cred.failedAttempts || 0) + 1;
  let lockedUntil = '';
  if (newCount >= MAX_FAILED_ATTEMPTS) {
    lockedUntil = fmt_(new Date(Date.now() + LOCKOUT_MINUTES * 60000), 'dd-MM-yyyy HH:mm:ss');
  }
  sheet.getRange(cred.rowNum, 6, 1, 2).setValues([[newCount, lockedUntil]]);
}

function resetFailedAttempts_(empId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('AdminCredentials');
  const cred = getAdminCredential_(empId);
  if (!cred) return;
  // col 5=Last Login (set to now), col 6=Failed Attempts (reset to 0), col 7=Locked Until (clear)
  sheet.getRange(cred.rowNum, 5, 1, 3).setValues([[fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'), 0, '']]);
}

function isLocked_(cred) {
  if (!cred || !cred.lockedUntil) return false;
  const t = parseDdMmYyyy_(cred.lockedUntil);
  return t && t > Date.now();
}

function parseDdMmYyyy_(s) {
  if (!s) return 0;
  const p = String(s).split(/[- :]/);
  if (p.length < 3) return 0;
  return new Date(p[2], p[1] - 1, p[0], p[3] || 0, p[4] || 0, p[5] || 0).getTime();
}



// ═══════════════════════════════════════════════════════════════════
//  SECTION 4 — SUBMIT / OUTCOME / DELETE
// ═══════════════════════════════════════════════════════════════════

function submitEntry(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const cfg      = getConfig();
    const verified = getHeadcountMap_()[String(payload.emp || '').trim()];
    if (!verified) throw new Error('Agent verification failed. Please re-login.');

    const caseNo = sanitize_(payload.caseNo);
    const cbDate = String(payload.cbDate || '').trim();
    const cbTime = String(payload.cbTime || '').trim();
    if (!caseNo) throw new Error('Case Number is required.');
    if (!cbDate) throw new Error('Callback Date is required.');
    if (!cbTime) throw new Error('Callback Time is required.');

    const scheduled = parseScheduled_(cbDate, cbTime);
    if (scheduled.getTime() < Date.now() - 2 * 60000) throw new Error('Callback time is in the past.');

    const todayCount = countEntriesToday_(verified.emp);
    if (todayCount >= Number(cfg.MAX_DAILY_ENTRIES_PER_AGENT)) {
      throw new Error('Daily entry limit reached (' + cfg.MAX_DAILY_ENTRIES_PER_AGENT + ').');
    }

    const dup = findDuplicate_(caseNo, Number(cfg.DUPLICATE_WINDOW_HOURS));
    if (dup && !payload.forceSubmit) {
      return { ok: false, duplicate: true,
        message: 'Case #' + caseNo + ' was already logged on ' + dup.timestamp + ' by ' + dup.agent + '. Submit anyway?' };
    }

    const entryId  = Utilities.getUuid().substring(0, 8).toUpperCase();
    const timestamp = fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss');
    const rowData  = [
      entryId, timestamp, verified.emp, verified.name, verified.supervisor, verified.manager,
      caseNo, cbDate, payload.priority || 'Medium', cbTime, sanitize_(payload.reason), sanitize_(payload.notes),
      verified.advisorEmail, 'Pending', 'FALSE', '', '', '', scheduled.toISOString()
    ];

    const primary = SpreadsheetApp.getActive().getSheetByName(cfg.CMS_SHEET);
    primary.appendRow(rowData);

    if (secondaryBreakerOpen_()) {
      queueSyncRetry_(rowData);
    } else {
      try {
        const ss2    = SpreadsheetApp.openById(cfg.SECONDARY_SHEET_ID);
        const sheet2 = ss2.getSheetByName(cfg.CMS_SHEET) || ss2.getSheets()[0];
        sheet2.appendRow(rowData);
        resetSecondaryBreaker_();
      } catch (err) {
        trackSecondaryFailure_();
        queueSyncRetry_(rowData);
        logError_('submitEntry.secondarySync', err.message, verified.emp);
      }
    }
    

    audit_('SUBMIT', verified.emp, entryId, 'Case ' + caseNo + ' @ ' + cbTime);
    lock.releaseLock();
    return { ok: true, entryId: entryId, message: 'Case logged successfully.',
      todayCount: todayCount + 1, upcoming: getUpcomingForAgent_(verified.emp, 10) };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('submitEntry', e.message, (payload && payload.emp) || '');
    throw e;
  }
}

function markOutcome(empId, entryId, outcome, outcomeNotes, reschedTime) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(entryId) && String(data[i][2]) === String(empId)) {
        applyOutcome_(sheet, i + 1, data[i], outcome, outcomeNotes, reschedTime, empId);
        lock.releaseLock();
        return { ok: true, upcoming: getUpcomingForAgent_(empId, 10) };
      }
    }
    throw new Error('Entry not found or not owned by you.');
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('markOutcome', e.message, empId);
    throw e;
  }
}

/**
 * Core outcome writer. Called by both agent and admin paths.
 * Also sends Completed and Rescheduled emails to TL/Manager.
 */
function applyOutcome_(sheet, rowNum, origRow, outcome, outcomeNotes, reschedTime, actingEmp) {
  const status      = (outcome === 'Rescheduled') ? 'Rescheduled' : 'Completed';
  const agentName   = String(origRow[3]);
  const agentEmp    = String(origRow[2]);
  const caseNo      = String(origRow[6]);
  const cbDate      = String(origRow[7]);
  const cbTime      = String(origRow[9]);
  const priority    = String(origRow[8]);
  const reason      = String(origRow[10]);
  const newDate     = reschedTime ? reschedTime.date : '';
  const newTime     = reschedTime ? reschedTime.time : '';

  sheet.getRange(rowNum, 14).setValue(status);
  sheet.getRange(rowNum, 16).setValue(fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'));
  sheet.getRange(rowNum, 17).setValue(outcome);
  sheet.getRange(rowNum, 18).setValue(sanitize_(outcomeNotes || ''));
  audit_('OUTCOME:' + outcome, actingEmp, origRow[0], outcomeNotes || '');

  // Create new rescheduled entry
  if (outcome === 'Rescheduled' && reschedTime) {
    const newEntryId = Utilities.getUuid().substring(0, 8).toUpperCase();
    const newSched   = parseScheduled_(reschedTime.date, reschedTime.time);
    sheet.appendRow([
      newEntryId, fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'), origRow[2], origRow[3], origRow[4], origRow[5],
      origRow[6], reschedTime.date, origRow[8], reschedTime.time, origRow[10],
      '\u21BB Rescheduled from ' + origRow[0] + '. ' + sanitize_(outcomeNotes || ''),
      origRow[12], 'Pending', 'FALSE', '', '', '', newSched.toISOString()
    ]);
    audit_('RESCHEDULE_NEW', actingEmp, newEntryId, 'From ' + origRow[0]);
  }

  // Send email to TL (To) and Manager (CC)
  var chain = getTlAndManager_(agentEmp);
  var tl    = chain.tl;
  var mgr   = chain.manager;

  if (outcome === 'Completed' && (tl || mgr)) {
    var isOthers = outcomeNotes && outcomeNotes.length > 0;
    sendEmailSafe_({
      to:      tl  ? tl.advisorEmail  : (mgr ? mgr.advisorEmail : ''),
      cc:      mgr ? mgr.advisorEmail : '',
      subject: '\u2705 Case Completed \u2014 #' + caseNo + ' by ' + agentName,
      html: buildEmailHtml_({
        headerColor:    '#3B6D11',
        title:          '\u2705 Case Completed \u2014 #' + caseNo,
        tagline:        'A callback in your team has been marked as completed.',
        greeting:       'Hi ' + (tl ? tl.name.split(' ')[0] : 'Team') + ',',
        alertMsg:       isOthers
          ? '<b>Closed via Others:</b> ' + sanitize_(outcomeNotes)
          : '<b>Case #' + caseNo + ' has been completed</b> by ' + agentName + '.',
        alertColor:     '#EAF3DE', alertBorder: '#3B6D11', alertTextColor: '#27500A',
        rows: [
          ['Agent',        agentName + ' (' + agentEmp + ')'],
          ['Case Number',  '#' + caseNo],
          ['Completed At', fmt_(new Date(), 'dd-MM-yyyy HH:mm') + ' IST'],
          ['Priority',     priority],
          ['Reason',       reason],
          ['Outcome',      isOthers ? 'Closed via Others \u2014 ' + sanitize_(outcomeNotes) : 'Completed'],
          ['TL (To)',      tl  ? tl.name  : '\u2014'],
          ['Manager (CC)', mgr ? mgr.name : '\u2014'],
        ],
        cta: 'View Entry in CMS',
      })
    });
  }

  if (outcome === 'Rescheduled' && (tl || mgr)) {
    sendEmailSafe_({
      to:      tl  ? tl.advisorEmail  : (mgr ? mgr.advisorEmail : ''),
      cc:      mgr ? mgr.advisorEmail : '',
      subject: '\u21BB Case Rescheduled \u2014 #' + caseNo + ' \u2192 ' + newDate + ' at ' + newTime,
      html: buildEmailHtml_({
        headerColor:    '#534AB7',
        title:          '\u21BB Case Rescheduled \u2014 #' + caseNo,
        tagline:        'A callback has been rescheduled. A new entry has been created.',
        greeting:       'Hi ' + (tl ? tl.name.split(' ')[0] : 'Team') + ',',
        alertMsg:       '<b>Case #' + caseNo + ' has been rescheduled</b> from ' + cbTime +
                        ' to ' + newDate + ' at ' + newTime + ' IST. Original entry closed.',
        alertColor:     '#EEEDFE', alertBorder: '#534AB7', alertTextColor: '#3C3489',
        rows: [
          ['Agent',         agentName + ' (' + agentEmp + ')'],
          ['Case Number',   '#' + caseNo],
          ['Original Time', cbDate + ' at ' + cbTime + ' IST'],
          ['New Time',      newDate + ' at ' + newTime + ' IST'],
          ['Priority',      priority],
          ['Reason',        reason],
          ['Note',          outcomeNotes || 'No Answer \u2014 auto-rescheduled'],
        ],
        cta: 'View New Entry in CMS',
      })
    });
  }
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 5 — ADMIN ACTIONS
// ═══════════════════════════════════════════════════════════════════

function inScope_(actor, row) {
  if (actor.role === 'WFM') return true;
  if (actor.role === 'TM')  return row[5] === actor.name;
  if (actor.role === 'TL')  return row[4] === actor.name;
  return false;
}

function adminMarkOutcome(actorEmp, entryId, outcome, outcomeNotes, reschedTime) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const lock  = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(entryId)) {
        if (!inScope_(actor, data[i])) throw new Error('Entry is outside your scope.');
        applyOutcome_(sheet, i + 1, data[i], outcome, outcomeNotes, reschedTime, actorEmp);
        lock.releaseLock();
        return { ok: true, dashboard: buildScopedDashboard_(actor) };
      }
    }
    throw new Error('Entry not found.');
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminMarkOutcome', e.message, actorEmp);
    throw e;
  }
}

function adminEditEntry(actorEmp, entryId, edits) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const lock  = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(entryId)) {
        if (!inScope_(actor, data[i])) throw new Error('Entry is outside your scope.');
        const rowNum = i + 1;
        if (edits.caseNo  !== undefined) sheet.getRange(rowNum, 7).setValue(sanitize_(edits.caseNo));
        if (edits.cbDate)                sheet.getRange(rowNum, 8).setValue(edits.cbDate);
        if (edits.priority)              sheet.getRange(rowNum, 9).setValue(edits.priority);
        if (edits.cbTime)                sheet.getRange(rowNum, 10).setValue(edits.cbTime);
        if (edits.reason)                sheet.getRange(rowNum, 11).setValue(sanitize_(edits.reason));
        if (edits.notes  !== undefined)  sheet.getRange(rowNum, 12).setValue(sanitize_(edits.notes));
        if (edits.cbDate || edits.cbTime) {
          const newD = edits.cbDate || data[i][7];
          const newT = edits.cbTime || data[i][9];
          sheet.getRange(rowNum, 19).setValue(parseScheduled_(newD, newT).toISOString());
          sheet.getRange(rowNum, 15).setValue('FALSE');
        }
        audit_('ADMIN_EDIT', actorEmp, entryId, JSON.stringify(edits));
        lock.releaseLock();
        return { ok: true, dashboard: buildScopedDashboard_(actor) };
      }
    }
    throw new Error('Entry not found.');
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminEditEntry', e.message, actorEmp);
    throw e;
  }
}

function adminDeleteEntry(actorEmp, entryId, reason) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const lock  = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(entryId)) {
        if (!inScope_(actor, data[i])) throw new Error('Entry is outside your scope.');
        audit_('ADMIN_DELETE', actorEmp, entryId,
          'agent=' + data[i][2] + ' case=' + data[i][6] + ' reason=' + (reason || '(none)'));
        sheet.deleteRow(i + 1);
        lock.releaseLock();
        return { ok: true, dashboard: buildScopedDashboard_(actor) };
      }
    }
    throw new Error('Entry not found.');
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminDeleteEntry', e.message, actorEmp);
    throw e;
  }
}

function adminReassignEntry(actorEmp, entryId, newAgentEmp, note) {
  const actor    = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const newAgent = getHeadcountMap_()[String(newAgentEmp || '').trim()];
  if (!newAgent) throw new Error('New agent not found.');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(entryId)) {
        if (!inScope_(actor, data[i])) throw new Error('Entry is outside your scope.');
        const rowNum = i + 1;
        sheet.getRange(rowNum, 3, 1, 4).setValues([[newAgent.emp, newAgent.name, newAgent.supervisor, newAgent.manager]]);
        sheet.getRange(rowNum, 13).setValue(newAgent.advisorEmail);
        sheet.getRange(rowNum, 15).setValue('FALSE');
        audit_('ADMIN_REASSIGN', actorEmp, entryId, 'from=' + data[i][2] + ' to=' + newAgent.emp);
        lock.releaseLock();
        return { ok: true, dashboard: buildScopedDashboard_(actor) };
      }
    }
    throw new Error('Entry not found.');
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminReassignEntry', e.message, actorEmp);
    throw e;
  }
}

function adminNudge(actorEmp, entryId, message) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(entryId)) {
      if (!inScope_(actor, data[i])) throw new Error('Entry is outside your scope.');
      const agentEmail = data[i][12];
      if (!agentEmail) throw new Error('Target agent has no email on file.');
      sendEmailSafe_({
        to:      agentEmail,
        cc:      actor.advisorEmail,
        subject: '\uD83D\uDCE3 Nudge from ' + actor.name + ' \u2014 Case #' + data[i][6],
        body:    actor.name + ' (' + actor.role + ') sent a nudge:\n\n' +
                 '  Case No : ' + data[i][6] + '\n  Time: ' + data[i][9] + ' on ' + data[i][7] + '\n\n' +
                 'Message: ' + (message || '(none)') + '\n\n\u2014 CMS v3.6'
      });
      audit_('ADMIN_NUDGE', actorEmp, entryId, 'target=' + data[i][2]);
      return { ok: true };
    }
  }
  throw new Error('Entry not found.');
}

function adminBulkMarkOutcome(actorEmp, entryIds, outcome, outcomeNotes) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  if (!Array.isArray(entryIds) || !entryIds.length) throw new Error('No entries provided.');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet  = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data   = sheet.getDataRange().getValues();
    const idSet  = new Set(entryIds.map(String));
    let done = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (!idSet.has(String(data[i][0]))) continue;
      if (!inScope_(actor, data[i]) || data[i][13] !== 'Pending') continue;
      applyOutcome_(sheet, i + 1, data[i], outcome, outcomeNotes, null, actorEmp);
      done++;
    }
    audit_('ADMIN_BULK_OUTCOME:' + outcome, actorEmp, '', 'count=' + done);
    lock.releaseLock();
    return { ok: true, done: done, dashboard: buildScopedDashboard_(actor) };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminBulkMarkOutcome', e.message, actorEmp);
    throw e;
  }
}

function adminBulkReassign(actorEmp, entryIds, newAgentEmp) {
  const actor    = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const newAgent = getHeadcountMap_()[String(newAgentEmp || '').trim()];
  if (!newAgent) throw new Error('New agent not found.');
  if (!Array.isArray(entryIds) || !entryIds.length) throw new Error('No entries provided.');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    const idSet = new Set(entryIds.map(String));
    let done = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (!idSet.has(String(data[i][0]))) continue;
      if (!inScope_(actor, data[i]) || data[i][13] !== 'Pending') continue;
      const rowNum = i + 1;
      sheet.getRange(rowNum, 3, 1, 4).setValues([[newAgent.emp, newAgent.name, newAgent.supervisor, newAgent.manager]]);
      sheet.getRange(rowNum, 13).setValue(newAgent.advisorEmail);
      sheet.getRange(rowNum, 15).setValue('FALSE');
      done++;
    }
    audit_('ADMIN_BULK_REASSIGN', actorEmp, '', 'count=' + done + ' to=' + newAgentEmp);
    lock.releaseLock();
    return { ok: true, done: done, dashboard: buildScopedDashboard_(actor) };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('adminBulkReassign', e.message, actorEmp);
    throw e;
  }
}



// ═══════════════════════════════════════════════════════════════════
//  SECTION 6 — DASHBOARDS
// ═══════════════════════════════════════════════════════════════════

function buildScopedDashboard_(actor, daysFilter) {
  const cfg   = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.CMS_SHEET);

  // Early exit on empty sheet — return zeroed stats
  if (!sheet || sheet.getLastRow() < 2) {
    return { stats: { pending: 0, completed: 0, overdue: 0, today: 0, noAnswer: 0 },
             agents: [], lobs: [], entries: [], partial: false, daysFilter: 30,
             queryMs: 0, repeatOffenders: [] };
  }

  const days       = Math.min(Number(daysFilter) || Number(cfg.DASHBOARD_DAYS_DEFAULT) || 30, Number(cfg.DASHBOARD_DAYS_MAX) || 180);
  const cutoffMs   = Date.now() - days * 86400000;
  const queryStart = Date.now();
  let partial      = false;

  const totalRows = sheet.getLastRow();
  const readStart = Math.max(2, totalRows - DASHBOARD_MAX_ROWS); 
  const headerRow = sheet.getRange(1, 1, 1, 19).getValues()[0];
  const data      = [headerRow].concat(sheet.getRange(readStart, 1, totalRows - readStart + 1, 19).getValues());

  const now           = Date.now();
  const todayStr      = fmt_(new Date(), 'dd-MM-yyyy');
  const overdueCutoff = Number(cfg.OVERDUE_THRESHOLD_MIN) * 60000;
  const stats         = { pending: 0, completed: 0, overdue: 0, today: 0, noAnswer: 0 };
  const byAgent       = {};
  const byLOB         = {};
  const entries       = [];
  const map           = getHeadcountMap_();

  for (let i = 1; i < data.length; i++) {
    if (Date.now() - queryStart > DASHBOARD_QUERY_TIMEOUT_MS) {
      partial = true;
      logError_('buildScopedDashboard_.timeout', 'Stopped at row ' + i, actor.emp);
      break;
    }
    if (!inScope_(actor, data[i])) continue;

        let loggedMs = 0;
    if (data[i][1] instanceof Date) {
      loggedMs = data[i][1].getTime();
    } else {
      const lv = data[i][1];
      if (lv instanceof Date) {
        loggedMs = lv.getTime();
      } else {
        const p = String(lv).split(/[- :]/);
        if (p.length >= 3) loggedMs = new Date(p[2], p[1]-1, p[0], p[3]||0, p[4]||0, p[5]||0).getTime();
      }
    }
    if (!loggedMs || loggedMs < cutoffMs) continue;

    const status   = data[i][13];
    const agent    = data[i][3];
    const agentEmp = String(data[i][2]);
    const agentLOB = (map[agentEmp] && map[agentEmp].lob) || '\u2014';
    const sched    = data[i][18] ? new Date(data[i][18]).getTime() : 0;

    if (!byAgent[agent]) byAgent[agent] = { emp: agentEmp, name: agent, pending: 0, overdue: 0, completed: 0, noAnswer: 0 };
    if (!byLOB[agentLOB]) byLOB[agentLOB] = { lob: agentLOB, pending: 0, overdue: 0, completed: 0 };

    if (status === 'Pending') {
      stats.pending++; byAgent[agent].pending++; byLOB[agentLOB].pending++;
      if (sched && sched < now - overdueCutoff) {
        stats.overdue++; byAgent[agent].overdue++; byLOB[agentLOB].overdue++;
      }
    } else if (status === 'Completed') {
      stats.completed++; byAgent[agent].completed++; byLOB[agentLOB].completed++;
    }
    if (data[i][16] === 'NoAnswer') { stats.noAnswer++; byAgent[agent].noAnswer++; }
    const loggedDateStr = data[i][1] instanceof Date ? fmt_(data[i][1], 'dd-MM-yyyy') : String(data[i][1]).substring(0, 10);
if (loggedDateStr === todayStr) stats.today++;

    if (entries.length < DASHBOARD_MAX_ROWS) {
      entries.push({
        entryId: data[i][0], loggedAt: data[i][1], emp: agentEmp, agent: agent,
        supervisor: data[i][4], manager: data[i][5], caseNo: data[i][6],
        cbDate: fmtDateCell_(data[i][7]), priority: data[i][8], cbTime: fmtTimeCell_(data[i][9]),
        reason: data[i][10], notes: data[i][11], status: status,
        reminderSent: data[i][14] === 'TRUE', outcome: data[i][16], outcomeNotes: data[i][17],
        scheduled: sched, lob: agentLOB,
        overdue: status === 'Pending' && sched && sched < now - overdueCutoff
      });
    }
  }

  const safeEntries = entries.sort((a, b) => (b.scheduled || 0) - (a.scheduled || 0));

  // Repeat offender detection — WFM only
  let repeatOffenders = [];
  if (actor.role === 'WFM') {
    const caseMap = {};
    for (const e of safeEntries) {
      if (!e.caseNo) continue;
      if (!caseMap[e.caseNo]) caseMap[e.caseNo] = { count: 0, agents: new Set(), lastDate: '' };
      caseMap[e.caseNo].count++;
      caseMap[e.caseNo].agents.add(e.agent);
      if (!caseMap[e.caseNo].lastDate || e.cbDate > caseMap[e.caseNo].lastDate) caseMap[e.caseNo].lastDate = e.cbDate;
    }
    repeatOffenders = Object.entries(caseMap)
      .filter(([, v]) => v.count >= 3)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 10)
      .map(([caseNo, v]) => ({ caseNo, count: v.count, agents: [...v.agents], lastDate: v.lastDate }));
  }

  return {
    stats:    stats,
    agents:   Object.values(byAgent).sort((a, b) => (b.pending + b.overdue) - (a.pending + a.overdue)),
    lobs:     Object.values(byLOB).sort((a, b) => b.pending - a.pending),
    entries:  safeEntries,
    partial:  partial,
    daysFilter: days,
    queryMs:  Date.now() - queryStart,
    repeatOffenders: repeatOffenders
  };
}

function getTeamRoster(actorEmp) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  const map   = getHeadcountMap_();
  const roster = Object.values(map).filter(p => {
    if (actor.role === 'WFM') return p.role === 'Agent' || p.role === 'TL';
    if (actor.role === 'TM')  return p.manager === actor.name && (p.role === 'Agent' || p.role === 'TL');
    return (p.supervisor === actor.name && p.role === 'Agent') || p.emp === actor.emp;
  });
  return roster.map(p => ({ emp: p.emp, name: p.name, lob: p.lob, role: p.role }))
               .sort((a, b) => a.name.localeCompare(b.name));
}

function refreshDashboard(actorEmp, daysFilter) {
  const actor = assertRole_(actorEmp, ['TL', 'TM', 'WFM']);
  return buildScopedDashboard_(actor, daysFilter);
}

function refreshQADashboard(actorEmp, daysFilter) {
  const profile = getHeadcountMap_()[String(actorEmp || '').trim()];
  if (!profile || profile.role !== 'QA') throw new Error('QA role required.');
  audit_('QA_DASHBOARD_PULL', actorEmp, '', 'days=' + (daysFilter || 30));
  const wfmScope = { ...profile, role: 'WFM' };
  return buildScopedDashboard_(wfmScope, daysFilter);
}

function qaExportCSV(actorEmp, fromDate, toDate) {
  const profile = getHeadcountMap_()[String(actorEmp || '').trim()];
  if (!profile || profile.role !== 'QA') throw new Error('QA role required.');
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  const data  = sheet.getDataRange().getValues();
  const from  = fromDate ? new Date(fromDate + 'T00:00:00').getTime() : 0;
  const to    = toDate   ? new Date(toDate   + 'T23:59:59').getTime() : Number.MAX_SAFE_INTEGER;
  const SAFE_HEADERS = ['Entry ID','Logged At','Emp ID','Agent','Supervisor','Manager','Case No','Callback Date','Priority','Callback Time','Reason','Notes','Status','Outcome'];
  const rows = [SAFE_HEADERS];
  for (let i = 1; i < data.length; i++) {
    const parts  = String(data[i][1]).split(/[- :]/);
    const logged = new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0).getTime();
    if (logged < from || logged > to) continue;
    rows.push([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],
               data[i][6],data[i][7],data[i][8],data[i][9],data[i][10],data[i][11],data[i][13],data[i][16]]);
  }
  audit_('QA_EXPORT', actorEmp, '', 'rows='+(rows.length-1));
  return rows.map(r => r.map(c => {
    const s = String(c == null ? '' : c).replace(/"/g, '""');
    return /[",\n]/.test(s) ? '"'+s+'"' : s;
  }).join(',')).join('\n');
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 7 — WFM CONFIG / EXPORT / AUDIT
// ═══════════════════════════════════════════════════════════════════

function getConfigForAdmin_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) return [];
  const data  = sheet.getDataRange().getValues();
  const out   = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) out.push({ key: String(data[i][0]), value: String(data[i][1] == null ? '' : data[i][1]), notes: String(data[i][2] || '') });
  }
  return out;
}

function adminUpdateConfig(actorEmp, key, value) {
  assertRole_(actorEmp, ['WFM']);
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(key).trim()) {
      sheet.getRange(i + 1, 2).setValue(value);
      refreshConfigCache();
      audit_('CONFIG_UPDATE', actorEmp, '', key + '=' + value);
      return { ok: true, config: getConfigForAdmin_() };
    }
  }
  sheet.appendRow([key, value, '']);
  refreshConfigCache();
  audit_('CONFIG_ADD', actorEmp, '', key + '=' + value);
  return { ok: true, config: getConfigForAdmin_() };
}

function adminExportCSV(actorEmp, fromDate, toDate) {
  const actor = assertRole_(actorEmp, ['TM', 'WFM']);
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  const data  = sheet.getDataRange().getValues();
  const from  = fromDate ? new Date(fromDate + 'T00:00:00').getTime() : 0;
  const to    = toDate   ? new Date(toDate   + 'T23:59:59').getTime() : Number.MAX_SAFE_INTEGER;
  const rows  = [data[0]];
  for (let i = 1; i < data.length; i++) {
    if (!inScope_(actor, data[i])) continue;
    const parts  = String(data[i][1]).split(/[- :]/);
    const logged = new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0).getTime();
    if (logged < from || logged > to) continue;
    rows.push(data[i]);
  }
  audit_('ADMIN_EXPORT', actorEmp, '', 'rows='+(rows.length-1));
  return rows.map(r => r.map(c => {
    const s = String(c == null ? '' : c).replace(/"/g, '""');
    return /[",\n]/.test(s) ? '"'+s+'"' : s;
  }).join(',')).join('\n');
}

function adminGetAuditTail(actorEmp, limit) {
  assertRole_(actorEmp, ['WFM']);
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().AUDIT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const n     = Math.min(limit || 50, 200);
  const start = Math.max(2, sheet.getLastRow() - n + 1);
  const data  = sheet.getRange(start, 1, sheet.getLastRow() - start + 1, 6).getValues();
  return data.reverse().map(r => ({ ts: r[0], action: r[1], empId: r[2], entryId: r[3], details: r[4], user: r[5] }));
}

function adminImpersonate(actorEmp, targetEmp) {
  assertRole_(actorEmp, ['WFM']);
  const target = getHeadcountMap_()[String(targetEmp || '').trim()];
  if (!target) throw new Error('Target agent not found.');
  audit_('ADMIN_IMPERSONATE', actorEmp, '', 'target=' + targetEmp);
  return { ...target, upcoming: getUpcomingForAgent_(targetEmp, 10), todayCount: countEntriesToday_(targetEmp), impersonated: true };
}



// ═══════════════════════════════════════════════════════════════════
//  SECTION 8 — AGENT QUERIES + STATS
// ═══════════════════════════════════════════════════════════════════

function getUpcomingForAgent_(empId, limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data  = sheet.getDataRange().getValues();
  const now   = Date.now();
  const out   = [];
  for (let i = data.length - 1; i >= 1 && out.length < (limit || 10); i--) {
    if (String(data[i][2]) !== String(empId)) continue;
    if (data[i][13] !== 'Pending') continue;
    let sched = 0;
    if (data[i][18]) {
      const sv = data[i][18];
      sched = sv instanceof Date ? sv.getTime() : new Date(String(sv)).getTime();
      if (isNaN(sched)) sched = 0;
    }
    if (!sched) sched = parseScheduled_(data[i][7], data[i][9]).getTime();
    out.push({
      entryId: data[i][0], caseNo: data[i][6],
      cbDate: fmtDateCell_(data[i][7]), cbTime: fmtTimeCell_(data[i][9]),
      priority: data[i][8], reason: data[i][10], notes: data[i][11],
      scheduled: sched, minsUntil: Math.round((sched - now) / 60000), overdue: sched < now
    });
  }
  return out.sort((a, b) => a.scheduled - b.scheduled);
}
function searchMyHistory(empId, query) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data  = sheet.getDataRange().getValues();
  const q     = String(query || '').toLowerCase().trim();
  const out   = [];
  for (let i = data.length - 1; i >= 1 && out.length < 25; i--) {
    if (String(data[i][2]) !== String(empId)) continue;
    const hay = (String(data[i][6]) + ' ' + String(data[i][10]) + ' ' + String(data[i][11])).toLowerCase();
    if (q && hay.indexOf(q) === -1) continue;
    out.push({ entryId: data[i][0], caseNo: data[i][6], cbDate: fmtDateCell_(data[i][7]),
      cbTime: String(data[i][9]), priority: data[i][8], status: data[i][13], outcome: data[i][16], notes: data[i][11] });
  }
  return out;
}

/**
 * Returns performance stats for an agent.
 * Week runs Sunday (0) → Saturday (6).
 */
function getAgentStats(actorEmp) {
  const profile = getHeadcountMap_()[String(actorEmp || '').trim()];
  if (!profile) throw new Error('Employee not found.');

  const sheet   = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  const data    = sheet.getDataRange().getValues();
  const now     = new Date();
  const cutoff30 = new Date(now.getTime() - 30 * 86400000);

  // Week: Sunday=0 ... Saturday=6
  const dayOfWeek = now.getDay();
  const weekStart = new Date(now.getTime() - dayOfWeek * 86400000);
  weekStart.setHours(0, 0, 0, 0);
  const weekEnd = new Date(weekStart.getTime() + 7 * 86400000);

  let completed = 0, noAnswer = 0, pending = 0, total = 0, thisWeek = 0, today = 0;
  let totalResponseMins = 0, responseCount = 0;
  const daily    = [0, 0, 0, 0, 0, 0, 0]; // Sun..Sat
  const todayStr = fmt_(now, 'dd-MM-yyyy');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== String(actorEmp)) continue;
    
    // REPLACED: Old split pattern with parseLoggedAt_
    const loggedMs = parseLoggedAt_(data[i][1]);
    if (!loggedMs || loggedMs < cutoff30.getTime()) continue;
    const logged = new Date(loggedMs);
    
    total++;
    const status = String(data[i][13] || '');
    if (status === 'Completed') completed++;
    else if (status === 'NoAnswer' || status === 'No Answer') noAnswer++;
    else pending++;
    
    if (String(data[i][1]).startsWith(todayStr)) today++;
    if (logged >= weekStart && logged < weekEnd) {
      thisWeek++;
      daily[logged.getDay()]++;
    }
    
    if (status === 'Completed' && data[i][7] && data[i][9]) {
      try {
        // Also fix the callback time parsing using the helper
        const loggedMsCb = parseLoggedAt_(data[i][1]);
        const cbDate   = String(data[i][7]).split('-');
        const cbTime   = String(data[i][9]).split(':');
        const schedMs  = new Date(cbDate[0], cbDate[1]-1, cbDate[2], cbTime[0]||0, cbTime[1]||0).getTime();
        const diffMins = Math.abs(schedMs - loggedMsCb) / 60000;
        if (diffMins < 1440) { totalResponseMins += diffMins; responseCount++; }
      } catch(e) {}
    }
  }

  const weeklyTrend = daily.map(function(count, i) { return { count: count, isToday: i === dayOfWeek }; });
  audit_('AGENT_STATS', actorEmp, '', 'total='+total+' week='+thisWeek);
  return { total, completed, noAnswer, pending, thisWeek, today,
    avgResponseMins: responseCount > 0 ? totalResponseMins / responseCount : 0, weeklyTrend };
}

function getPendingReassignsForAgent_(empId) {
  const sheet  = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  const data   = sheet.getDataRange().getValues();
  const hcMap  = getHeadcountMap_();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== String(empId)) continue;
    if (String(data[i][13]) !== 'Pending_Ack') continue;
    const supProfile = hcMap[String(data[i][4] || '')] || {};
    results.push({ entryId: data[i][0], caseNo: data[i][6], cbDate: data[i][7], cbTime: data[i][9],
      priority: data[i][8], reason: data[i][10], lob: data[i][12] || '', note: data[i][17] || '', fromName: supProfile.name || '' });
  }
  return results;
}

function agentAckReassign(actorEmp, entryId, accept, declineReason) {
  const profile = getHeadcountMap_()[String(actorEmp || '').trim()];
  if (!profile) throw new Error('Employee not found.');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
    const data  = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entryId) && String(data[i][2]) === String(actorEmp)) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) throw new Error('Entry not found or not assigned to you.');
    if (accept) {
      if (String(data[targetRow - 1][13]) === 'Pending_Ack') sheet.getRange(targetRow, 14).setValue('Pending');
      audit_('AGENT_ACK_ACCEPT', actorEmp, entryId, '');
      lock.releaseLock();
      return { ok: true, upcoming: getUpcomingForAgent_(actorEmp, 10) };
    } else {
      const orig = String(data[targetRow - 1][11] || '');
      sheet.getRange(targetRow, 12).setValue(orig + (orig ? ' | ' : '') + 'DECLINED: ' + String(declineReason || '').trim());
      sheet.getRange(targetRow, 14).setValue('Declined_Reassign');
      audit_('AGENT_ACK_DECLINE', actorEmp, entryId, declineReason || '');
      lock.releaseLock();
      return { ok: true };
    }
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('agentAckReassign', e.message, actorEmp);
    throw e;
  }
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 9 — QA
// ═══════════════════════════════════════════════════════════════════

function qaFlagEntry(actorEmp, entryId, reason) {
  const profile = getHeadcountMap_()[String(actorEmp || '').trim()];
  if (!profile || profile.role !== 'QA') throw new Error('QA role required.');
  if (!entryId) throw new Error('Entry ID required.');
  const sheet = SpreadsheetApp.getActive().getSheetByName('Flags');
  if (!sheet) throw new Error('Flags sheet missing. Run setupCMSv3().');
  const cleanReason = String(reason || '').trim();
  const data = sheet.getLastRow() >= 2 ? sheet.getDataRange().getValues() : [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entryId)) {
      if (!cleanReason) { sheet.deleteRow(i + 1); audit_('QA_FLAG_REMOVED', actorEmp, entryId, ''); return { ok: true }; }
      sheet.getRange(i + 1, 3, 1, 3).setValues([[cleanReason, actorEmp, fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss')]]);
      audit_('QA_FLAG_UPDATED', actorEmp, entryId, cleanReason);
      return { ok: true };
    }
  }
  if (cleanReason) {
    sheet.appendRow([entryId, fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'), cleanReason, actorEmp, fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss')]);
    audit_('QA_FLAG_SET', actorEmp, entryId, cleanReason);
  }
  return { ok: true };
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 10 — TRIGGERS + EMAIL SWEEPS
// ═══════════════════════════════════════════════════════════════════

function installAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReminders_').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('sendEscalations_').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('retryQueueSweep').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(21).everyDays(1).create();
  ScriptApp.newTrigger('dailyHealthCheck').timeBased().atHour(6).everyDays(1).create();
  SpreadsheetApp.getActive().toast('\u2705 All CMS triggers installed (5 active).', 'CMS v3.6', 5);
}

/**
 * EMAIL TYPE 1: Reminder — fires 15 min and 5 min before due.
 * TL in To, Manager in CC.
 */
function sendReminders_() {
  const cfg   = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data  = sheet.getDataRange().getValues();
  const now   = Date.now();
  const win5  = Number(cfg.REMINDER_WINDOW_MIN) || 5;
  const win15 = Number(cfg.REMINDER_EARLY_MIN)  || 15;
  const tol   = Number(cfg.REMINDER_TOLERANCE_MIN) || 1.5;
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][13] !== 'Pending' || !data[i][18]) continue;
    const sched     = new Date(data[i][18]).getTime();
    const diffMins  = (sched - now) / 60000;
    const agentEmp  = String(data[i][2]);
    const agentName = String(data[i][3]);
    const caseNo    = String(data[i][6]);
    const cbDate    = String(data[i][7]);
    const cbTime    = fmtTimeCell_(data[i][9]);
    const priority  = String(data[i][8]);
    const reason    = String(data[i][10]);
    const notes     = String(data[i][11] || '');
    const chain     = getTlAndManager_(agentEmp);
    const tl        = chain.tl;
    const manager   = chain.manager;

    // 15-minute reminder
    if (diffMins >= win15 - tol && diffMins <= win15 + tol && data[i][14] !== 'EARLY_SENT') {
      sendEmailSafe_({
        to:      tl      ? tl.advisorEmail      : data[i][12],
        cc:      manager ? manager.advisorEmail  : '',
        subject: '\u23F0 Callback Reminder \u2014 ' + agentName + ': Case #' + caseNo + ' at ' + cbTime + ' IST',
        html: buildEmailHtml_({
          headerColor:   '#1D4ED8',
          title:         '\u23F0 15-Minute Callback Reminder',
          tagline:       'A callback assigned to your agent is due in 15 minutes.',
          greeting:      'Hi ' + (tl ? tl.name.split(' ')[0] : 'Team') + ',',
          alertMsg:      '<b>15-minute reminder:</b> Case #' + caseNo + ' is due at <b>' + cbTime +
                         ' IST</b>. Agent: <b>' + agentName + '</b>.',
          alertColor:    '#E6F1FB', alertBorder: '#185FA5', alertTextColor: '#0C447C',
          rows: [
            ['Agent',        agentName + ' (' + agentEmp + ')'],
            ['Case Number',  '#' + caseNo],
            ['Scheduled',    cbDate + ' at ' + cbTime + ' IST'],
            ['Priority',     priority],
            ['Reason',       reason],
            ['Notes',        notes || '\u2014'],
            ['TL (To)',      tl      ? tl.name      : '\u2014'],
            ['Manager (CC)', manager ? manager.name  : '\u2014'],
          ],
          cta:  'View in CMS Portal',
          note: 'A 5-minute reminder will also fire before the call.',
        })
      });
      sheet.getRange(i + 1, 15).setValue('EARLY_SENT');
      audit_('REMINDER_15M', agentEmp, data[i][0], 'tl=' + (tl ? tl.advisorEmail : '—'));
    }

    // 5-minute reminder
    if (diffMins >= win5 - tol && diffMins <= win5 + tol && data[i][14] !== 'TRUE') {
      sendEmailSafe_({
        to:      tl      ? tl.advisorEmail      : data[i][12],
        cc:      manager ? manager.advisorEmail  : '',
        subject: '\u23F0 Callback in 5 minutes \u2014 ' + agentName + ': Case #' + caseNo + ' at ' + cbTime + ' IST',
        html: buildEmailHtml_({
          headerColor:   '#185FA5',
          title:         '\u23F0 5-Minute Callback Reminder',
          tagline:       'A callback assigned to your agent is due in 5 minutes.',
          greeting:      'Hi ' + (tl ? tl.name.split(' ')[0] : 'Team') + ',',
          alertMsg:      '<b>5-minute reminder:</b> Case #' + caseNo + ' is due at <b>' + cbTime +
                         ' IST</b>. Agent: <b>' + agentName + '</b>. If not actioned on time, an L1 escalation will fire.',
          alertColor:    '#E6F1FB', alertBorder: '#185FA5', alertTextColor: '#0C447C',
          rows: [
            ['Agent',        agentName + ' (' + agentEmp + ')'],
            ['Case Number',  '#' + caseNo],
            ['Scheduled',    cbDate + ' at ' + cbTime + ' IST'],
            ['Priority',     priority],
            ['Reason',       reason],
            ['TL (To)',      tl      ? tl.name      : '\u2014'],
            ['Manager (CC)', manager ? manager.name  : '\u2014'],
          ],
          cta: 'View in CMS Portal',
        })
      });
      updates.push(i + 1);
      audit_('REMINDER_5M', agentEmp, data[i][0], 'tl=' + (tl ? tl.advisorEmail : '—'));
    }
  }
  if (updates.length) updates.forEach(r => sheet.getRange(r, 15).setValue('TRUE'));
}

/**
 * EMAIL TYPES 2 & 3: L1 (5 min overdue) and L2 (10 min overdue) escalations.
 * L1: TL + Manager both in To.
 * L2: Manager in To, TL in CC.
 * Uses Script Properties to avoid duplicate sends.
 */
function sendEscalations_() {
  const cfg   = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data  = sheet.getDataRange().getValues();
  const now   = Date.now();
  const props = PropertiesService.getScriptProperties();
  const sent  = JSON.parse(props.getProperty('cms_escalated') || '{}');
  let changed = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][13]) !== 'Pending' || !data[i][18]) continue;
    const scheduled  = new Date(data[i][18]).getTime();
    const overdueMs  = now - scheduled;
    if (overdueMs <= 0) continue;

    const entryId    = String(data[i][0]);
    const agentEmp   = String(data[i][2]);
    const agentName  = String(data[i][3]);
    const caseNo     = String(data[i][6]);
    const cbDate     = String(data[i][7]);
    const cbTime     = fmtTimeCell_(data[i][9]);
    const priority   = String(data[i][8]);
    const reason     = String(data[i][10]);
    const notes      = String(data[i][11] || '');
    const overdueMin = Math.round(overdueMs / 60000);
    const chain      = getTlAndManager_(agentEmp);
    const tl         = chain.tl;
    const manager    = chain.manager;
    const tlEmail    = tl      ? tl.advisorEmail      : '';
    const mgrEmail   = manager ? manager.advisorEmail  : '';
    const tlName     = tl      ? tl.name.split(' ')[0] : 'Team';
    const mgrName    = manager ? manager.name.split(' ')[0] : 'Manager';

    // L1 — 5 minutes overdue
    if (overdueMin >= 5 && !sent[entryId + '_L1']) {
      const toAddrs = [tlEmail, mgrEmail].filter(Boolean);
      if (toAddrs.length) {
        sendEmailSafe_({
          to:      toAddrs,
          cc:      '',
          subject: '\uD83D\uDD34 L1 Escalation \u2014 Case #' + caseNo + ' overdue by ' + overdueMin + ' minutes',
          html: buildEmailHtml_({
            headerColor:   '#EA7317',
            title:         '\uD83D\uDD34 L1 Escalation \u2014 ' + overdueMin + ' minutes overdue',
            tagline:       'A callback has passed its scheduled time. Immediate action required.',
            greeting:      'Hi ' + tlName + (mgrName !== 'Manager' ? ' & ' + mgrName : '') + ',',
            alertMsg:      '<b>L1 Escalation:</b> Case #' + caseNo + ' was due at <b>' + cbTime +
                           ' IST</b> and is now <b>' + overdueMin + ' minutes overdue</b>. ' +
                           'Please nudge or reassign immediately. An L2 escalation fires in 5 minutes if still pending.',
            alertColor:    '#FAEEDA', alertBorder: '#EA7317', alertTextColor: '#633806',
            rows: [
              ['Agent',        agentName + ' (' + agentEmp + ')'],
              ['Case Number',  '#' + caseNo],
              ['Scheduled At', cbDate + ' at ' + cbTime + ' IST'],
              ['Overdue By',   overdueMin + ' minutes'],
              ['Priority',     priority],
              ['Reason',       reason],
              ['Notes',        notes || '\u2014'],
            ],
            cta:  'Nudge or Reassign in CMS',
            note: 'L2 escalation fires in 5 minutes with Manager in To.',
          })
        });
      }
      sent[entryId + '_L1'] = true;
      changed = true;
      audit_('L1_ESCALATION', agentEmp, entryId, 'overdueMin=' + overdueMin);
    }

    // L2 — 10 minutes overdue
    if (overdueMin >= 10 && sent[entryId + '_L1'] && !sent[entryId + '_L2']) {
      if (mgrEmail) {
        sendEmailSafe_({
          to:      mgrEmail,
          cc:      tlEmail,
          subject: '\uD83D\uDEA8 L2 Escalation \u2014 Case #' + caseNo + ' overdue ' + overdueMin + ' min \u2014 Immediate action required',
          html: buildEmailHtml_({
            headerColor:   '#A32D2D',
            title:         '\uD83D\uDEA8 L2 Escalation \u2014 ' + overdueMin + ' minutes overdue',
            tagline:       'Manager-level action required. TL has been CC\'d.',
            greeting:      'Hi ' + mgrName + ',',
            alertMsg:      '<b>L2 Escalation:</b> Case #' + caseNo + ' is now <b>' + overdueMin +
                           ' minutes overdue</b>. The L1 escalation was sent 5 minutes ago. ' +
                           'Please take ownership or direct your TL to act immediately.',
            alertColor:    '#FCEBEB', alertBorder: '#A32D2D', alertTextColor: '#791F1F',
            rows: [
              ['Agent',        agentName + ' (' + agentEmp + ')'],
              ['Case Number',  '#' + caseNo],
              ['Scheduled At', cbDate + ' at ' + cbTime + ' IST'],
              ['Overdue By',   overdueMin + ' minutes'],
              ['Priority',     priority],
              ['Reason',       reason],
              ['Manager (To)', manager ? manager.name : '\u2014'],
              ['TL (CC)',      tl      ? tl.name      : '\u2014'],
              ['L1 Sent',      'Yes \u2014 at 5 minutes overdue'],
            ],
            cta: 'Escalate / Reassign in CMS',
          })
        });
      }
      sent[entryId + '_L2'] = true;
      changed = true;
      audit_('L2_ESCALATION', agentEmp, entryId, 'overdueMin=' + overdueMin);
    }
  }

  // Purge resolved entries from tracker
  const pendingIds = new Set();
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][13]) === 'Pending') pendingIds.add(String(data[j][0]));
  }
  Object.keys(sent).forEach(function(k) {
    const entId = k.replace('_L1', '').replace('_L2', '');
    if (!pendingIds.has(entId)) { delete sent[k]; changed = true; }
  });

  if (changed) props.setProperty('cms_escalated', JSON.stringify(sent));
}

function retryQueueSweep() {
  const props = PropertiesService.getScriptProperties();
  const q = JSON.parse(props.getProperty('SYNC_QUEUE') || '[]');
  if (!q.length) return;
  const cfg = getConfig();
  const remaining = []; let ok = 0;
  try {
    const ss2    = SpreadsheetApp.openById(cfg.SECONDARY_SHEET_ID);
    const sheet2 = ss2.getSheetByName(cfg.CMS_SHEET) || ss2.getSheets()[0];
    q.forEach(item => { try { sheet2.appendRow(item.row); ok++; } catch (e) { remaining.push(item); } });
  } catch (e) { logError_('retryQueueSweep', e.message, ''); return; }
  props.setProperty('SYNC_QUEUE', JSON.stringify(remaining));
  if (ok > 0) audit_('SYNC_RETRY', 'SYSTEM', '', 'Replayed ' + ok + ', pending ' + remaining.length);
}

function dailyDigest() {
  const cfg   = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(cfg.CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data  = sheet.getDataRange().getValues();
  const today = fmt_(new Date(), 'dd-MM-yyyy');
  const byTL  = {};
  
  for (let i = 1; i < data.length; i++) {
    const loggedVal = data[i][1];
    const loggedStr = loggedVal instanceof Date ? fmt_(loggedVal, 'dd-MM-yyyy') : String(loggedVal).substring(0, 10);
    if (loggedStr !== today) continue;  // Only count today's entries
    
    const tl = data[i][4];
    if (!byTL[tl]) byTL[tl] = { total: 0, completed: 0, pending: 0, noAnswer: 0, byAgent: {} };
    byTL[tl].total++;
    if (data[i][13] === 'Completed') byTL[tl].completed++;
    else if (data[i][13] === 'Pending') byTL[tl].pending++;
    if (data[i][16] === 'NoAnswer') byTL[tl].noAnswer++;
    byTL[tl].byAgent[data[i][3]] = (byTL[tl].byAgent[data[i][3]] || 0) + 1;
  }
  
  const map = getHeadcountMap_();
  Object.entries(byTL).forEach(([tlName, s]) => {
    const tl = Object.values(map).find(h => h.name === tlName && ['TL','TM','WFM'].includes(h.role));
    if (!tl || !tl.advisorEmail) return;
    const top = Object.entries(s.byAgent).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([n, c]) => '  \u2022 ' + n + ': ' + c).join('\n');
    sendEmailSafe_({
      to:      tl.advisorEmail,
      subject: '\uD83D\uDCCA CMS Daily Digest \u2014 ' + today,
      body:    'CMS v3.6 Daily Digest \u2014 ' + today + '\n\nTeam: ' + tlName + '\n\n' +
               '  Total: ' + s.total + ' | Completed: ' + s.completed + ' | Pending: ' + s.pending +
               ' | No-Answer: ' + s.noAnswer + '\n  Completion: ' +
               (s.total ? Math.round(100 * s.completed / s.total) : 0) + '%\n\nTop agents:\n' + top + '\n\n\u2014 CMS v3.6'
    });
  });
}
function dailyHealthCheck() {
  try {
    const cfg     = getConfig();
    const props   = PropertiesService.getScriptProperties();
    const syncQ   = JSON.parse(props.getProperty('SYNC_QUEUE') || '[]');
    const breaker = secondaryBreakerOpen_();
    const hcMap   = getHeadcountMap_();
    const wfmEmails = Object.values(hcMap).filter(p => p.role === 'WFM' && p.advisorEmail).map(p => p.advisorEmail);
    if (!wfmEmails.length) return;
    const flags = [];
    if (syncQ.length > 100) flags.push('\u26A0 SYNC QUEUE BACKED UP: ' + syncQ.length);
    if (breaker) flags.push('\u26A0 SECONDARY SHEET CIRCUIT BREAKER OPEN');
    const status = flags.length === 0 ? '\u2705 All systems green' : '\u26A0\uFE0F ' + flags.length + ' issue(s)';
    sendEmailSafe_({
      to:      wfmEmails.join(','),
      subject: 'CMS Daily Health \u2014 ' + fmt_(new Date(), 'dd-MM-yyyy') + ' \u2014 ' + status,
      body:    'CMS v3.6 daily health report\n\n' + status + '\n\nSync queue: ' + syncQ.length +
               '\nCircuit breaker: ' + (breaker ? 'OPEN' : 'closed') + '\n\n\u2014 CMS v3.6'
    });
    audit_('HEALTH_CHECK', 'SYSTEM', '', status);
  } catch (e) { logError_('dailyHealthCheck', e.message, ''); }
}



// ═══════════════════════════════════════════════════════════════════
//  SECTION 11 — EMAIL: RELAY + HTML BUILDER
// ═══════════════════════════════════════════════════════════════════

/**
 * Sends email through Account B's relay web app.
 * opts.to and opts.cc can be a string or array of strings.
 */
function sendEmailSafe_(opts) {
  if (!opts || !opts.to || !opts.subject) return;
  var toStr  = Array.isArray(opts.to) ? opts.to.filter(Boolean).join(',') : String(opts.to  || '');
  var ccStr  = Array.isArray(opts.cc) ? opts.cc.filter(Boolean).join(',') : String(opts.cc  || '');
  if (!toStr) return;

  // Separate HTML and plain-text body correctly.
  // The relay script must receive 'htmlBody' for HTML emails.
  var hasHtml   = !!(opts.html);
  var htmlBody  = hasHtml ? opts.html  : '';
  var plainBody = opts.body || (hasHtml ? 'Please view this email in an HTML-capable client.' : '');

  try {
    var payload = {
      secret:    EMAIL_RELAY_SECRET,
      to:        toStr,
      cc:        ccStr,
      subject:   opts.subject,
      plainBody: plainBody,
      htmlBody:  htmlBody,      // relay uses this for GmailApp htmlBody option
    };

    var response = UrlFetchApp.fetch(EMAIL_RELAY_URL, {
      method:             'post',
      contentType:        'application/json',
      muteHttpExceptions: true,
      payload:            JSON.stringify(payload),
    });

    var code = response.getResponseCode();
    var text = response.getContentText();

    if (code !== 200) {
      logError_('sendEmailSafe_', 'Relay HTTP ' + code + ': ' + text.substring(0, 200), toStr);
      return;
    }

    try {
      var result = JSON.parse(text);
      if (!result.ok) {
        logError_('sendEmailSafe_', 'Relay error: ' + (result.error || text), toStr);
      }
    } catch(parseErr) {
      logError_('sendEmailSafe_', 'Relay bad JSON: ' + text.substring(0, 200), toStr);
    }

  } catch(e) {
    logError_('sendEmailSafe_', 'Relay fetch failed: ' + e.message, toStr);
  }
}

function buildEmailHtml_(opts) {
  var hdrColor  = opts.headerColor    || '#FF385C';
  var alertBg   = opts.alertColor     || '#E6F1FB';
  var alertBdr  = opts.alertBorder    || '#185FA5';
  var alertTxt  = opts.alertTextColor || '#0C447C';
  var ctaUrl    = opts.ctaUrl         || String(getConfig().PORTAL_URL || 'https://script.google.com/');
  var rowsHtml  = (opts.rows || []).map(function(r) {
    return '<tr><td style="padding:9px 13px;background:#f2f2f2;color:#888;font-weight:600;font-size:13px;width:160px;border-bottom:1px solid #eee;font-family:Arial">' + r[0] + '</td>' +
           '<td style="padding:9px 13px;color:#1a1a1a;font-size:13px;border-bottom:1px solid #eee;font-family:Arial">' + r[1] + '</td></tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:20px;background:#f5f5f5;font-family:Arial,sans-serif">' +
    '<div style="max-width:580px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden">' +
    '<div style="background:' + hdrColor + ';padding:16px 20px;display:flex;align-items:center;gap:10px">' +
      '<div style="width:34px;height:34px;border-radius:7px;background:rgba(255,255,255,.2);display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700;color:#fff;font-family:Arial">C</div>' +
      '<div><div style="font-size:15px;font-weight:700;color:#fff;font-family:Arial">CMS Enterprise Portal</div>' +
      '<div style="font-size:11px;color:rgba(255,255,255,.8);font-family:Arial">Teleperformance \u00B7 Workforce Management \u00B7 RGS Mumbai</div></div>' +
    '</div>' +
    '<div style="padding:20px">' +
      '<div style="font-size:18px;font-weight:700;color:#1a1a1a;margin-bottom:4px;font-family:Arial">' + opts.title + '</div>' +
      '<div style="font-size:13px;color:#666;margin-bottom:16px;font-family:Arial">' + (opts.tagline || '') + '</div>' +
      '<hr style="border:none;border-top:1px solid #eee;margin:12px 0">' +
      '<p style="font-size:14px;color:#333;margin-bottom:14px;font-family:Arial">' + (opts.greeting || 'Hi,') + '</p>' +
      '<div style="padding:11px 14px;border-radius:6px;background:' + alertBg + ';border-left:4px solid ' + alertBdr + ';color:' + alertTxt + ';font-size:13px;margin-bottom:16px;line-height:1.55;font-family:Arial">' + opts.alertMsg + '</div>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #e8e8e8;border-radius:6px;overflow:hidden;margin-bottom:18px">' + rowsHtml + '</table>' +
      '<div style="text-align:center;margin-bottom:18px"><a href="' + ctaUrl + '" style="display:inline-block;padding:11px 26px;background:' + hdrColor + ';color:#fff;border-radius:999px;text-decoration:none;font-size:13px;font-weight:700;font-family:Arial">' + (opts.cta || 'Open in CMS Portal') + '</a></div>' +
      '<hr style="border:none;border-top:1px solid #eee;margin:12px 0">' +
      (opts.note ? '<p style="font-size:12px;color:#aaa;text-align:center;line-height:1.6;font-family:Arial"><b>Note:</b> ' + opts.note + '</p>' : '') +
    '</div>' +
    '<div style="background:#f2f2f2;padding:12px 20px;text-align:center;font-size:11px;color:#aaa;border-top:1px solid #e5e5e5;font-family:Arial">' +
      'CMS Enterprise Portal v3.6 \u00B7 Teleperformance WFM RTA, RGS Mumbai \u00B7 Built by Jackson Koli<br>' +
      '<span style="color:#ccc">Sent via CMS Email Relay</span>' +
    '</div></div></body></html>';
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 12 — NOTIFICATIONS
// ═══════════════════════════════════════════════════════════════════

function sendNotification(senderEmp, targetType, targetValue, title, body) {
  const sender  = assertRole_(senderEmp, ['Agent','TL','TM','WFM']);
  title  = sanitize_(String(title  || '').substring(0, 100));
  body   = sanitize_(String(body   || '').substring(0, 500));
  if (!title && !body) throw new Error('Please enter a message.');
  const scoping = validateNotificationScope_(sender, targetType, targetValue);
  if (!scoping.ok) throw new Error(scoping.reason);
  const sheet   = SpreadsheetApp.getActive().getSheetByName('Notifications');
  if (!sheet) throw new Error('Notifications sheet missing. Run setupCMSv3().');
  const notifId = Utilities.getUuid().substring(0, 8).toUpperCase();
  sheet.appendRow([notifId, fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'),
    sender.emp, sender.name, sender.role, targetType, targetValue, title, body, '']);
  audit_('NOTIFICATION_SENT', sender.emp, notifId, 'target=' + targetType + ':' + targetValue);
  return { ok: true, notifId: notifId, recipientCount: scoping.recipientCount };
}

function validateNotificationScope_(sender, targetType, targetValue) {
  const map    = getHeadcountMap_();
  const people = Object.values(map);
  if (sender.role === 'Agent') {
    if (targetType !== 'emp') return { ok: false, reason: 'Agents can only ping their supervisor.' };
    const target = map[String(targetValue).trim()];
    if (!target || target.name !== sender.supervisor) return { ok: false, reason: 'Agents can only ping their own supervisor.' };
    return { ok: true, recipientCount: 1 };
  }
  if (sender.role === 'TL') {
    if (targetType === 'team' && targetValue === sender.name) {
      return { ok: true, recipientCount: people.filter(p => p.supervisor === sender.name && p.role === 'Agent').length };
    }
    if (targetType === 'emp') {
      const target = map[String(targetValue).trim()];
      if (!target) return { ok: false, reason: 'Target not found.' };
      if (target.supervisor === sender.name || target.name === sender.manager) return { ok: true, recipientCount: 1 };
      return { ok: false, reason: 'TL can only message own team or manager.' };
    }
    return { ok: false, reason: 'TL can message own team or a specific member.' };
  }
  if (sender.role === 'TM') {
    if (targetType === 'manager' && targetValue === sender.name) return { ok: true, recipientCount: people.filter(p => p.manager === sender.name && p.role === 'TL').length };
    if (targetType === 'lob') {
      if (sender.lob !== targetValue) return { ok: false, reason: 'TM can only broadcast to own LOB.' };
      return { ok: true, recipientCount: people.filter(p => p.lob === targetValue && p.manager === sender.name).length };
    }
    if (targetType === 'emp') {
      const target = map[String(targetValue).trim()];
      if (!target) return { ok: false, reason: 'Target not found.' };
      if (target.manager === sender.name || target.supervisor === sender.name) return { ok: true, recipientCount: 1 };
      return { ok: false, reason: 'TM can only message people in their chain.' };
    }
    return { ok: false, reason: 'Invalid target for TM.' };
  }
  if (sender.role === 'WFM') {
    if (targetType === 'all') return { ok: true, recipientCount: people.length };
    if (targetType === 'role') return { ok: true, recipientCount: people.filter(p => p.role === targetValue).length };
    if (targetType === 'lob')  return { ok: true, recipientCount: people.filter(p => p.lob === targetValue).length };
    if (targetType === 'emp')  { const t = map[String(targetValue).trim()]; return t ? { ok: true, recipientCount: 1 } : { ok: false, reason: 'Target not found.' }; }
    if (targetType === 'team')    return { ok: true, recipientCount: people.filter(p => p.supervisor === targetValue).length };
    if (targetType === 'manager') return { ok: true, recipientCount: people.filter(p => p.manager   === targetValue).length };
  }
  return { ok: false, reason: 'Unknown target type.' };
}

function pollNotifications(empId, sinceId) {
  try {
    const emp     = String(empId || '').trim();
    if (!emp) return [];
    const profile = getHeadcountMap_()[emp];
    if (!profile) return [];
    const sheet   = SpreadsheetApp.getActive().getSheetByName('Notifications');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const totalRows = sheet.getLastRow();
    const scanStart = Math.max(2, totalRows - 200);
    const data      = sheet.getRange(scanStart, 1, totalRows - scanStart + 1, 10).getValues();
    const cutoff    = Date.now() - 86400000;
    const unread    = [];
    for (let i = 0; i < data.length; i++) {
      const notifId = String(data[i][0]);
      if (!notifId) continue;
      if (sinceId && notifId === sinceId) break;
      const p = String(data[i][1]).split(/[- :]/);
      if (p.length >= 3 && new Date(p[2], p[1]-1, p[0], p[3]||0, p[4]||0, p[5]||0).getTime() < cutoff) continue;
      const readBy = String(data[i][9] || '').split(',');
      if (readBy.indexOf(emp) !== -1) continue;
      if (!notificationTargetsUser_(profile, data[i][5], data[i][6])) continue;
      unread.push({ notifId, createdAt: data[i][1], senderEmp: data[i][2], senderName: data[i][3],
        senderRole: data[i][4], targetType: data[i][5], targetValue: data[i][6], title: data[i][7], body: data[i][8] });
    }
    return unread.reverse();
  } catch (e) { logError_('pollNotifications', e.message, empId); return []; }
}

function notificationTargetsUser_(profile, targetType, targetValue) {
  if (targetType === 'all') return true;
  if (targetType === 'role')    return profile.role       === targetValue;
  if (targetType === 'lob')     return profile.lob        === targetValue;
  if (targetType === 'emp')     return profile.emp        === String(targetValue).trim();
  if (targetType === 'team')    return profile.supervisor === targetValue;
  if (targetType === 'manager') return profile.manager    === targetValue;
  return false;
}

function markNotificationRead(empId, notifId) {
  try {
    const emp   = String(empId || '').trim();
    if (!emp || !notifId) return { ok: false };
    const sheet = SpreadsheetApp.getActive().getSheetByName('Notifications');
    if (!sheet) return { ok: false };
    const totalRows = sheet.getLastRow();
    const scanStart = Math.max(2, totalRows - 200);
    const data      = sheet.getRange(scanStart, 1, totalRows - scanStart + 1, 10).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(notifId)) {
        const rowNum = scanStart + i;
        const readBy = String(data[i][9] || '').split(',').filter(s => s);
        if (readBy.indexOf(emp) === -1) { readBy.push(emp); sheet.getRange(rowNum, 10).setValue(readBy.join(',')); }
        return { ok: true };
      }
    }
    return { ok: false };
  } catch (e) { logError_('markNotificationRead', e.message, empId); return { ok: false }; }
}

function getNotificationTargets(actorEmp) {
  const actor  = assertRole_(actorEmp, ['Agent','TL','TM','WFM']);
  const map    = getHeadcountMap_();
  const people = Object.values(map);
  const lobs   = [...new Set(people.map(p => p.lob).filter(l => l && l !== '\u2014'))].sort();
  const payload = { role: actor.role, name: actor.name, emp: actor.emp };
  if (actor.role === 'Agent') {
    const tl = people.find(p => p.name === actor.supervisor);
    payload.canSendTo = tl ? [{ type: 'emp', value: tl.emp, label: 'My TL: ' + tl.name }] : [];
  }
  if (actor.role === 'TL') {
    const myAgents = people.filter(p => p.supervisor === actor.name && p.role === 'Agent');
    payload.canSendTo = [{ type: 'team', value: actor.name, label: 'My team (' + myAgents.length + ')' }];
    const myTM = people.find(p => p.name === actor.manager);
    if (myTM) payload.canSendTo.push({ type: 'emp', value: myTM.emp, label: 'My TM: ' + myTM.name });
    payload.teamMembers = myAgents.map(a => ({ emp: a.emp, name: a.name }));
  }
  if (actor.role === 'TM') {
    const myTLs = people.filter(p => p.manager === actor.name && p.role === 'TL');
    payload.canSendTo = [
      { type: 'manager', value: actor.name, label: 'My TLs (' + myTLs.length + ')' },
      { type: 'lob', value: actor.lob, label: 'My LOB: ' + actor.lob }
    ];
    payload.chainMembers = people.filter(p => p.manager === actor.name || p.supervisor === actor.name)
      .map(p => ({ emp: p.emp, name: p.name, role: p.role }));
  }
  if (actor.role === 'WFM') {
    payload.canSendTo = [
      { type: 'all', value: 'all', label: 'Entire org (' + people.length + ')' },
      { type: 'role', value: 'Agent', label: 'All Agents' },
      { type: 'role', value: 'TL',    label: 'All TLs' },
      { type: 'role', value: 'TM',    label: 'All TMs' },
      { type: 'role', value: 'WFM',   label: 'All WFM' }
    ].concat(lobs.map(lob => ({ type: 'lob', value: lob, label: 'LOB: ' + lob })));
    payload.allPeople = people.map(p => ({ emp: p.emp, name: p.name, role: p.role, lob: p.lob }));
  }
  return payload;
}


// ═══════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
//  SECTION 8b — EARLY RESOLUTION
// ═══════════════════════════════════════════════════════════════════

function searchPendingByCaseNo(agentEmp, caseNo) {
  if (!agentEmp || !caseNo) throw new Error('Missing parameters.');
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, notFound: true, message: 'CMS sheet empty.' };
  const data = sheet.getDataRange().getValues();
  const needle = String(caseNo).trim().toLowerCase();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][6]).trim().toLowerCase() === needle && String(data[i][13]) === 'Pending') {
      return {
        ok: true, entryId: String(data[i][0]), caseNo: String(data[i][6]),
        origAgent: String(data[i][3]), origAgentEmp: String(data[i][2]),
        origDate: fmtDateCell_(data[i][7]), origTime: fmtTimeCell_(data[i][9]),
        priority: String(data[i][8]), reason: String(data[i][10]),
      };
    }
  }
  return { ok: false, notFound: true, message: 'No pending entry found for Case #' + caseNo };
}

function resolveEarlyByAnotherAgent(agentBEmp, caseNo, notes) {
  agentBEmp = String(agentBEmp || '').trim();
  caseNo    = sanitize_(String(caseNo || '').trim());
  if (!caseNo) throw new Error('Case number is required.');
  const hcMap  = getHeadcountMap_();
  const agentB = hcMap[agentBEmp];
  if (!agentB) throw new Error('Agent not found.');
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet) throw new Error('CMS sheet not found.');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = sheet.getDataRange().getValues();
    const now  = new Date();
    const nowFmt = fmt_(now, 'dd-MM-yyyy HH:mm:ss');
    let origRow = null, origRowNum = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][6]).trim().toLowerCase() === caseNo.toLowerCase() && String(data[i][13]) === 'Pending') {
        origRow = data[i]; origRowNum = i + 1; break;
      }
    }
    if (!origRow) {
      lock.releaseLock();
      return { ok: false, notFound: true, message: 'No pending entry found for Case #' + caseNo + '.' };
    }
    const origAgent    = String(origRow[3]);
    const origAgentEmp = String(origRow[2]);
    const origTime     = fmtTimeCell_(origRow[9]);
    const origDate     = fmtDateCell_(origRow[7]);
    const origEntryId  = String(origRow[0]);
    const priority     = String(origRow[8]);
    const reason       = String(origRow[10]);
    // Step 1: Close original as Resolved_Early
    const origNote = 'Resolved early by ' + agentB.name + ' (' + agentBEmp + ') at ' +
                     fmt_(now, 'HH:mm') + ' IST — customer called in before scheduled time.' +
                     (notes ? ' Note: ' + notes : '');
    sheet.getRange(origRowNum, 14).setValue('Resolved_Early');
    sheet.getRange(origRowNum, 16).setValue(nowFmt);
    sheet.getRange(origRowNum, 17).setValue('Resolved_Early');
    sheet.getRange(origRowNum, 18).setValue(origNote);
    // Step 2: New Completed entry under Agent B
    const newEntryId = Utilities.getUuid().substring(0, 8).toUpperCase();
    const newNote = '↙ Early resolution — Original Entry ' + origEntryId +
                    ' scheduled for ' + origDate + ' at ' + origTime + ' under ' + origAgent + '.' +
                    (notes ? ' Note: ' + notes : '');
    const scheduled = parseScheduled_(origRow[7], origRow[9]);
    sheet.appendRow([
      newEntryId, nowFmt, agentBEmp, agentB.name, agentB.supervisor, agentB.manager,
      origRow[6], origDate, priority, origTime, reason, newNote, agentB.advisorEmail,
      'Completed', 'FALSE', nowFmt, 'Resolved_Early', newNote, scheduled.toISOString()
    ]);
    // Step 3: Email original TL
    const chain = getTlAndManager_(origAgentEmp);
    const tl = chain.tl, mgr = chain.manager;
    if (tl || mgr) {
      sendEmailSafe_({
        to:      tl  ? tl.advisorEmail  : (mgr ? mgr.advisorEmail : ''),
        cc:      mgr ? mgr.advisorEmail : '',
        subject: 'ℹ️ Early Resolution — Case #' + origRow[6] + ' handled by ' + agentB.name,
        html: buildEmailHtml_({
          headerColor: '#008489',
          title:    '↙ Customer Called In Early — #' + origRow[6],
          tagline:  'This callback was resolved before the scheduled time. No action needed.',
          greeting: 'Hi ' + (tl ? tl.name.split(' ')[0] : 'Team') + ',',
          alertMsg: 'Case #' + origRow[6] + ' was scheduled for <b>' + origDate + ' at ' + origTime +
                    ' IST</b> under <b>' + origAgent + '</b>. The customer called in early and was handled by <b>' +
                    agentB.name + '</b> at <b>' + fmt_(now, 'HH:mm') + ' IST</b>. ' +
                    'The original entry has been automatically closed — no action needed.',
          alertColor: '#E8F4F5', alertBorder: '#008489', alertTextColor: '#004D50',
          rows: [
            ['Original Agent',    origAgent + ' (' + origAgentEmp + ')'],
            ['Handled By',        agentB.name + ' (' + agentBEmp + ')'],
            ['Case Number',       '#' + origRow[6]],
            ['Scheduled Time',    origDate + ' at ' + origTime + ' IST'],
            ['Resolved At',       fmt_(now, 'dd-MM-yyyy HH:mm') + ' IST'],
            ['Original Entry ID', origEntryId + ' — now Resolved_Early'],
            ['New Entry ID',      newEntryId + ' — under ' + agentB.name],
          ],
          cta:  'View in CMS Portal',
          note: origAgent + "'s entry is now closed. No escalation will fire.",
        })
      });
    }
    audit_('EARLY_RESOLUTION', agentBEmp, newEntryId, 'orig=' + origEntryId + ' origAgent=' + origAgentEmp);
    lock.releaseLock();
    return {
      ok: true, newEntryId, origEntryId, origAgent, origTime,
      upcoming: getUpcomingForAgent_(agentBEmp, 10),
      message: 'Entry ' + origEntryId + ' (' + origAgent + ') closed. New entry ' + newEntryId + ' created.',
    };
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    logError_('resolveEarlyByAnotherAgent', e.message, agentBEmp);
    throw e;
  }
}


//  SECTION 13 — HELPERS
// ═══════════════════════════════════════════════════════════════════

function sanitize_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function fmt_(d, pattern) { return Utilities.formatDate(d, getConfig().TIMEZONE, pattern); }
function fmtDateCell_(v)  { if (v instanceof Date) return fmt_(v, 'yyyy-MM-dd'); return String(v); }
function fmtTimeCell_(v)  {
  if (v instanceof Date) return fmt_(v, 'HH:mm');
  const s = String(v == null ? '' : v).trim();
  if (/^\d{1,2}:\d{2}:\d{2}$/.test(s)) return s.substring(0, 5);
  return s;
}

function parseScheduled_(dateVal, timeVal) {
  let y, mo, d;
  if (dateVal instanceof Date) { y = dateVal.getFullYear(); mo = dateVal.getMonth() + 1; d = dateVal.getDate(); }
  else { const p = String(dateVal).split('-').map(Number); y = p[0]; mo = p[1]; d = p[2]; }
  const t = parseTimeFlexible_(timeVal);
  return new Date(y, mo - 1, d, t.h, t.m, 0);
}

function parseTimeFlexible_(t) {
  if (t instanceof Date) return { h: t.getHours(), m: t.getMinutes() };
  const s = String(t).trim();
  if (!s) return { h: 0, m: 0 };
  if (s.toUpperCase().indexOf('M') !== -1) {
    const [main, mod] = s.split(' ');
    let [h, m] = main.split(':').map(Number);
    if (h === 12) h = 0;
    if (mod && mod.toUpperCase() === 'PM') h += 12;
    return { h, m: m || 0 };
  }
  const [h, m] = s.split(':').map(Number);
  return { h: h || 0, m: m || 0 };
}

function countEntriesToday_(empId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const data  = sheet.getDataRange().getValues();
  const today = fmt_(new Date(), 'dd-MM-yyyy');
  let n = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    // FIX: handle Date object
    const loggedStr = data[i][1] instanceof Date
      ? fmt_(data[i][1], 'dd-MM-yyyy')
      : String(data[i][1]).substring(0, 10);
    if (loggedStr !== today) continue;  // removed the `break` — rows may not be perfectly ordered
    if (String(data[i][2]) === String(empId)) n++;
  }
  return n;
}
function parseLoggedAt_(val) {
  if (val instanceof Date) return val.getTime();
  const p = String(val).split(/[- :]/);
  if (p.length >= 3) {
    // Handle 2-digit vs 4-digit years
    const year = p[2].length === 2 ? 2000 + parseInt(p[2], 10) : parseInt(p[2], 10);
    return new Date(year, parseInt(p[1], 10) - 1, parseInt(p[0], 10), 
                    parseInt(p[3] || 0, 10), parseInt(p[4] || 0, 10), parseInt(p[5] || 0, 10)).getTime();
  }
  return 0;
}
// Updated findDuplicate_
function findDuplicate_(caseNo, windowHours) {
  const sheet  = SpreadsheetApp.getActive().getSheetByName(getConfig().CMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data   = sheet.getDataRange().getValues();
  const cutoff = Date.now() - windowHours * 3600000;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][6]).trim() === String(caseNo).trim()) {
      // REPLACED: Old split pattern with parseLoggedAt_
      if (parseLoggedAt_(data[i][1]) >= cutoff) {
        return { timestamp: data[i][1], agent: data[i][3] };
      }
    }
  }
  return null;
}

function audit_(action, empId, entryId, details) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().AUDIT_SHEET);
    if (!sheet) return;
    sheet.appendRow([fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'), action, empId, entryId,
      sanitize_(details || ''), Session.getActiveUser().getEmail() || '']);
  } catch (e) {}
}

function logError_(fn, msg, context) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().ERROR_SHEET);
    if (sheet) sheet.appendRow([fmt_(new Date(), 'dd-MM-yyyy HH:mm:ss'), fn,
      sanitize_(msg), sanitize_(context || ''), Session.getActiveUser().getEmail() || '']);
    console.error('[' + fn + '] ' + msg);
  } catch (e) { console.error(e); }
}

function queueSyncRetry_(rowData) {
  const props = PropertiesService.getScriptProperties();
  const q     = JSON.parse(props.getProperty('SYNC_QUEUE') || '[]');
  q.push({ row: rowData, queuedAt: Date.now() });
  if (q.length > 5000) q.shift();
  props.setProperty('SYNC_QUEUE', JSON.stringify(q));
}

function secondaryBreakerOpen_() {
  const props    = PropertiesService.getScriptProperties();
  const openedAt = Number(props.getProperty('SEC_BREAKER_OPEN_AT') || 0);
  if (!openedAt) return false;
  const pauseMs  = (Number(getConfig().CIRCUIT_BREAKER_PAUSE_MIN) || 60) * 60000;
  if (Date.now() - openedAt > pauseMs) {
    props.deleteProperty('SEC_BREAKER_OPEN_AT');
    props.deleteProperty('SEC_BREAKER_FAIL_COUNT');
    return false;
  }
  return true;
}

function trackSecondaryFailure_() {
  const props     = PropertiesService.getScriptProperties();
  const threshold = Number(getConfig().CIRCUIT_BREAKER_THRESHOLD) || 3;
  const count     = Number(props.getProperty('SEC_BREAKER_FAIL_COUNT') || 0) + 1;
  props.setProperty('SEC_BREAKER_FAIL_COUNT', String(count));
  if (count >= threshold) {
    props.setProperty('SEC_BREAKER_OPEN_AT', String(Date.now()));
    audit_('CIRCUIT_BREAKER_OPEN', 'SYSTEM', '', 'Secondary sync paused');
  }
}

function resetSecondaryBreaker_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SEC_BREAKER_FAIL_COUNT')) {
    props.deleteProperty('SEC_BREAKER_FAIL_COUNT');
    props.deleteProperty('SEC_BREAKER_OPEN_AT');
  }
}


// ═══════════════════════════════════════════════════════════════════
//  SECTION 14 — SETUP + MENU
// ═══════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('\u26A1 CMS v3.6')
    .addItem('Setup / Repair sheets',   'setupCMSv3')
    .addItem('Install all triggers',    'installAllTriggers')
    .addItem('Show Web App URL',        'showWebAppUrl')
    .addSeparator()
    .addItem('Refresh caches',          'refreshConfigCache')
    .addItem('Run reminder sweep now',  'sendReminders_')
    .addItem('Run escalation sweep now','sendEscalations_')
    .addItem('Run health check now',    'dailyHealthCheck')
    .addToUi();
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('CMS Portal URL', url || 'Not deployed yet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupCMSv3() {
  const ss      = SpreadsheetApp.getActive();
  const headers = ['Entry ID','Logged At','Emp ID','Agent','Supervisor','Manager',
    'Case No','Callback Date','Priority','Callback Time','Reason','Notes',
    'Email','Status','Reminder Sent','Completed At','Outcome','Outcome Notes','Scheduled ISO'];
  ensureSheet_(ss, 'CMS',              headers);
  ensureSheet_(ss, 'AuditLog',         ['Timestamp','Action','Emp ID','Entry ID','Details','User']);
  ensureSheet_(ss, 'ErrorLog',         ['Timestamp','Function','Message','Context','User']);
  ensureSheet_(ss, 'AdminCredentials', ['Emp ID','Password Hash','Salt','Created','Last Login','Failed Attempts','Locked Until']);
  ensureSheet_(ss, 'Notifications',    ['Notif ID','Created At','Sender Emp','Sender Name','Sender Role','Target Type','Target Value','Title','Body','Read By']);
  ensureSheet_(ss, 'Flags',            ['Entry ID','Flagged At','Reason','QA Emp','Updated At']);
  const cfgSheet = ensureSheet_(ss, 'Config', ['Key','Value','Notes']);
  if (cfgSheet.getLastRow() < 2) {
    const rows = Object.entries(defaultConfig_()).map(([k, v]) => [k, v, '']);
    cfgSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  try {
    const cfg = getConfig();
    if (cfg.SECONDARY_SHEET_ID) {
      const ss2 = SpreadsheetApp.openById(cfg.SECONDARY_SHEET_ID);
      ensureSheet_(ss2, cfg.CMS_SHEET, headers);
    }
  } catch (e) { logError_('setupCMSv3.secondary', e.message, ''); }
  try {
    const credSheet = ss.getSheetByName('AdminCredentials');
    if (credSheet) {
      const p = credSheet.protect().setDescription('CMS: do not edit manually.');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    }
  } catch (e) {}
  SpreadsheetApp.getUi().alert('CMS v3.6 setup complete.\n\nNext steps:\n1. \u26A1 CMS v3.6 \u2192 Install all triggers\n2. Paste your EMAIL_RELAY_URL at the top of Code.gs\n3. Set PORTAL_URL in Config tab\n4. Deploy \u2192 New deployment \u2192 Web app');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0d1b3e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}


function testLogin(empId, password) {
  try {
    console.log('=== DEBUG testLogin ===');
    console.log('empId:', empId);  // FIXED: use parameter, not hardcoded value
    
    const profile = getHeadcountMap_()[String(empId).trim()];
    console.log('profile:', profile ? 'found' : 'not found');
    
    if (!profile) return { error: 'Profile not found' };
    
    const cred = getAdminCredential_(empId);
    console.log('cred:', cred ? 'exists' : 'not exists');
    
    return { 
      success: true, 
      profileExists: !!profile,
      role: profile ? profile.role : null,
      credExists: !!cred,
      firstTime: !cred,
      name: profile ? profile.name : null
    };
  } catch(e) {
    console.log('Error in testLogin:', e.message);
    return { error: e.message };
  }
}

function testDashboardOnly(actorEmp) {
  console.log('testDashboardOnly called for:', actorEmp);
  try {
    const actor = getHeadcountMap_()[String(actorEmp).trim()];
    if (!actor) {
      console.error('Actor not found');
      return { error: 'Actor not found' };
    }
    console.log('Actor role:', actor.role);
    console.log('Actor name:', actor.name);
    
    const dashboard = buildScopedDashboard_(actor, 30);
    console.log('Dashboard built, pending:', dashboard.stats.pending);
    
    return { 
      success: true, 
      pending: dashboard.stats.pending,
      completed: dashboard.stats.completed,
      overdue: dashboard.stats.overdue,
      agentCount: dashboard.agents.length
    };
  } catch(e) {
    console.error('testDashboardOnly error:', e.message);
    return { error: e.message, stack: e.stack };
  }
}
