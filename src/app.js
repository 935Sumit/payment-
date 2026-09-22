import { amountToWordsLine } from './words';
import { exportToExcel, parseExcelParties, exportMonthlyReport } from './excel';
import {
  getSupabaseConfig,
  saveSupabaseConfig,
  isSupabaseConfigured,
  testSupabaseConnection,
  getSupabaseSchemaSql,
  dbUpsertParties,
  dbDeleteParties,
  dbUpsertAccounts,
  dbDeleteAccount,
  dbUpsertHistories,
  dbDeleteHistory,
  dbSaveSetting,
  dbFetchSettings,
  dbPushAllToCloud,
  dbPullAllFromCloud
} from './supabase';

const STORE_KEYS = {
  parties: 'pv_parties',
  myAccounts: 'pv_my_accounts',
  history: 'pv_history',
  prefix: 'pv_words_prefix',
  route: 'pv_current_route',
  privacyMode: 'pv_privacy_mode',
  pin: 'pv_app_pin'
};

function loadStore(key, defaultValue = []) {
  try {
    const raw = localStorage.getItem(key);
    const val = raw ? JSON.parse(raw) : defaultValue;
    if (key === STORE_KEYS.parties && Array.isArray(val)) {
      // Auto-deduplicate by category and normalized account number/name
      const seen = new Set();
      const uniqueParties = [];
      val.forEach(p => {
        if (!p) return;
        const cat = p.category || 'material';
        const norm = (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
        const key = norm ? `${cat}_${norm}` : `${cat}_name_${(p.name || '').trim().toLowerCase()}`;
        if (!seen.has(key)) {
          seen.add(key);
          uniqueParties.push(p);
        }
      });
      return uniqueParties;
    }
    return val;
  } catch (e) {
    return defaultValue;
  }
}

function saveStore(key, data) {
  localStorage.setItem(key, JSON.stringify(data));
}

const VALID_ROUTES = ['dashboard', 'directory', 'newrun', 'history', 'accounts'];

function getRouteFromHash() {
  const hash = (window.location.hash || '').replace(/^#\/?/, '').split('?')[0].trim().toLowerCase();
  return VALID_ROUTES.includes(hash) ? hash : null;
}

let savedRoute = getRouteFromHash() || localStorage.getItem(STORE_KEYS.route) || 'dashboard';
if (!VALID_ROUTES.includes(savedRoute)) savedRoute = 'dashboard';

let state = {
  route: savedRoute,
  parties: loadStore(STORE_KEYS.parties),
  myAccounts: loadStore(STORE_KEYS.myAccounts),
  history: loadStore(STORE_KEYS.history),
  wordsPrefix: localStorage.getItem(STORE_KEYS.prefix) ?? 'INT ',
  privacyMode: localStorage.getItem(STORE_KEYS.privacyMode) !== 'false',
  unmaskedIds: new Set(),
  pin: localStorage.getItem(STORE_KEYS.pin) || null,
  isLocked: Boolean(localStorage.getItem(STORE_KEYS.pin)) && sessionStorage.getItem('pv_session_unlocked') !== 'true',
  pinInput: '',
  pinError: '',
  pinSuccess: false,
  directoryCategory: 'all',
  runCategory: 'all',
  run: {
    selectedIds: [],
    amounts: {},
    chequeNo: '',
    date: new Date().toISOString().slice(0, 10),
    accountId: '',
    prefix: localStorage.getItem(STORE_KEYS.prefix) ?? 'INT ',
    editingHistoryId: null // tracks if we are editing an existing run
  },
  supabase: {
    status: isSupabaseConfigured() ? 'checking' : 'unconfigured',
    message: '',
    isSyncing: false,
    lastSync: localStorage.getItem('pv_last_cloud_sync') || null,
    copiedSql: false
  },
  modal: null,
  search: '',
  historyPeriod: 'all',
  directorySelectedIds: [],
  quickAdd: {
    partyId: '',
    mode: 'directory',
    search: '',
  },
  monthlyExport: {
    month: new Date().getMonth() + 1,
    year: new Date().getFullYear()
  }
};

function navigateTo(route, pushHistory = true) {
  if (!VALID_ROUTES.includes(route)) route = 'dashboard';
  state.route = route;
  state.modal = null;
  state.search = '';
  state.directorySelectedIds = [];
  
  if (pushHistory) {
    if (window.location.hash !== `#/${route}`) {
      window.location.hash = `#/${route}`;
    }
  }
  render();
}

function openModal(modalObj, pushHistory = true) {
  state.modal = modalObj;
  if (pushHistory && !(history.state && history.state.isModal)) {
    history.pushState({ isModal: true, route: state.route }, '', `#/${state.route}`);
  }
  render();
}

function closeModal(fromPopState = false) {
  if (!state.modal) return;
  state.modal = null;
  resetQuickAdd();
  render();
  if (!fromPopState && history.state && history.state.isModal) {
    history.back();
  }
}

window.addEventListener('popstate', (e) => {
  // 1. If modal was open, user hit back to close modal
  if (state.modal) {
    state.modal = null;
    resetQuickAdd();
    render();
    return;
  }
  
  // 2. Otherwise update route based on hash or event state
  const hashRoute = getRouteFromHash();
  const targetRoute = hashRoute || (e.state && e.state.route) || 'dashboard';
  if (state.route !== targetRoute) {
    state.route = VALID_ROUTES.includes(targetRoute) ? targetRoute : 'dashboard';
    state.modal = null;
    state.search = '';
    state.directorySelectedIds = [];
    render();
  }
});

window.addEventListener('hashchange', () => {
  const hashRoute = getRouteFromHash();
  if (hashRoute && state.route !== hashRoute) {
    state.route = hashRoute;
    state.modal = null;
    state.search = '';
    state.directorySelectedIds = [];
    render();
  }
});

function uid() {
  return 'id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function persistParties() { 
  saveStore(STORE_KEYS.parties, state.parties);
  if (isSupabaseConfigured()) {
    dbUpsertParties(state.parties).catch(err => console.warn('Supabase auto-sync parties notice:', err));
  }
}

function removePartiesByIds(ids) {
  const idSet = new Set(ids);
  state.parties = state.parties.filter(p => !idSet.has(p.id));
  state.directorySelectedIds = state.directorySelectedIds.filter(pid => !idSet.has(pid));
  state.run.selectedIds = state.run.selectedIds.filter(pid => !idSet.has(pid));
  ids.forEach(pid => delete state.run.amounts[pid]);
  saveStore(STORE_KEYS.parties, state.parties);
  if (isSupabaseConfigured()) {
    dbDeleteParties(ids).catch(err => console.warn('Supabase auto-delete parties notice:', err));
  }
}

function persistAccounts() { 
  saveStore(STORE_KEYS.myAccounts, state.myAccounts);
  if (isSupabaseConfigured()) {
    dbUpsertAccounts(state.myAccounts).catch(err => console.warn('Supabase auto-sync accounts notice:', err));
  }
}

function persistHistory() { 
  saveStore(STORE_KEYS.history, state.history);
  if (isSupabaseConfigured()) {
    dbUpsertHistories(state.history).catch(err => console.warn('Supabase auto-sync history notice:', err));
  }
}

function persistPrefix() { 
  localStorage.setItem(STORE_KEYS.prefix, state.wordsPrefix);
  if (isSupabaseConfigured()) {
    dbSaveSetting('wordsPrefix', state.wordsPrefix).catch(err => console.warn('Supabase auto-sync prefix notice:', err));
  }
}

function formatINR(num) {
  num = Number(num) || 0;
  return num.toLocaleString('en-IN', { maximumFractionDigits: 0 });
}

function getChequeBookInfo(account) {
  if (!account || !account.chequeBookStart || !account.chequeBookEnd) {
    return null;
  }

  const startStr = String(account.chequeBookStart).trim();
  const endStr = String(account.chequeBookEnd).trim();
  const startNum = parseInt(startStr, 10);
  const endNum = parseInt(endStr, 10);

  if (isNaN(startNum) || isNaN(endNum) || startNum > endNum) {
    return null;
  }

  const totalLeaves = endNum - startNum + 1;
  const padLen = Math.max(startStr.length, endStr.length);

  // Find all used cheque numbers on this account from history
  const usedChequeSet = new Set();
  (state.history || []).forEach(h => {
    if (h.account && h.account.accountNo === account.accountNo && h.chequeNo) {
      usedChequeSet.add(String(h.chequeNo).trim());
    }
  });

  // Calculate used leaves in this series
  let usedCount = 0;
  for (let num = startNum; num <= endNum; num++) {
    const formatted = String(num).padStart(padLen, '0');
    const plain = String(num);
    if (usedChequeSet.has(formatted) || usedChequeSet.has(plain)) {
      usedCount++;
    }
  }

  // Find next unused cheque in series
  let nextUnused = null;
  for (let num = startNum; num <= endNum; num++) {
    const formatted = String(num).padStart(padLen, '0');
    const plain = String(num);
    if (!usedChequeSet.has(formatted) && !usedChequeSet.has(plain)) {
      nextUnused = formatted;
      break;
    }
  }

  const remainingLeaves = Math.max(0, totalLeaves - usedCount);
  const isOutOfLeaves = remainingLeaves === 0;
  const isNearEnd = remainingLeaves > 0 && remainingLeaves <= 3;

  return {
    startStr,
    endStr,
    startNum,
    endNum,
    padLen,
    totalLeaves,
    usedCount,
    remainingLeaves,
    nextUnused,
    isOutOfLeaves,
    isNearEnd
  };
}

function findDuplicateCheque(chequeNo, excludeHistoryId = null) {
  if (!chequeNo || !String(chequeNo).trim()) return null;
  const target = String(chequeNo).trim().toLowerCase();
  return (state.history || []).find(h => 
    h.id !== excludeHistoryId &&
    h.chequeNo &&
    String(h.chequeNo).trim().toLowerCase() === target
  );
}

const PARTY_CATEGORIES = [
  { id: 'material', label: 'Material & Vendor Bills', shortLabel: 'Material / Bills', icon: '📦' },
  { id: 'employees', label: 'Employees & Staff (Salary)', shortLabel: 'Employees / Staff', icon: '👤' },
  { id: 'family', legacyIds: ['contractors'], label: 'Family & Relatives', shortLabel: 'Family', icon: '👨‍👩‍👧‍👦' },
  { id: 'utilities', label: 'Rent, Utilities & Others', shortLabel: 'Rent / Utilities', icon: '🏢' }
];

function getCategoryMeta(catId) {
  return PARTY_CATEGORIES.find(c => c.id === catId || (c.legacyIds && c.legacyIds.includes(catId))) || PARTY_CATEGORIES[0];
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

function renderEmptyState({ icon = 'empty', title, message, buttons = [] }) {
  const btnHtml = buttons.length
    ? `<div class="empty-state-actions">${buttons.map(b => `
        <button type="button" class="btn ${b.variant || 'btn-primary'}" data-action="${b.action}">
          ${b.iconKey ? ICONS[b.iconKey] : ''}${b.iconKey ? ' ' : ''}${escapeHtml(b.label)}
        </button>
      `).join('')}</div>`
    : '';
  return `
    <div class="empty-state">
      <div class="glyph">${ICONS[icon] || ICONS.empty}</div>
      ${title ? `<h4>${escapeHtml(title)}</h4>` : ''}
      <p>${message}</p>
      ${btnHtml}
    </div>
  `;
}

const ICONS = {
  directory: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>`,
  newrun: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><line x1="9" y1="15" x2="15" y2="15"></line></svg>`,
  history: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>`,
  bank: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="10" width="18" height="12" rx="2"></rect><line x1="12" y1="1" x2="12" y2="10"></line><path d="M8 5h8"></path><path d="M3 10l9-7 9 7"></path></svg>`,
  plus: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>`,
  edit: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 1 1 3 3L12 15l-4 1 1-4z"></path></svg>`,
  trash: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>`,
  search: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>`,
  check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`,
  download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>`,
  upload: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>`,
  x: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`,
  empty: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`,
  mail: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>`,
  save: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>`,
  dashboard: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9"></rect><rect x="14" y="3" width="7" height="5"></rect><rect x="14" y="12" width="7" height="9"></rect><rect x="3" y="16" width="7" height="5"></rect></svg>`,
  cloud: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/></svg>`,
  database: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"></ellipse><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path></svg>`,
  copy: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>`,
  refresh: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>`,
  key: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21 2-2 2m-1.5 1.5L13 10m-3-1a5 5 0 1 0 0 10 5 5 0 0 0 0-10Zm0 0V2m7 5-2 2m0-2 2 2"></path></svg>`,
  uploadCloud: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 16 12 12 8 16"></polyline><line x1="12" y1="12" x2="12" y2="21"></line><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"></path><polyline points="16 16 12 12 8 16"></polyline></svg>`,
  downloadCloud: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 17 12 21 16 17"></polyline><line x1="12" y1="12" x2="12" y2="21"></line><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"></path><polyline points="8 17 12 21 16 17"></polyline></svg>`,
  eye: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"></path><circle cx="12" cy="12" r="3"></circle></svg>`,
  eyeOff: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9.88 9.88a3 3 0 1 0 4.24 4.24"></path><path d="M10.73 5.08A10.43 10.43 0 0 1 12 5c7 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"></path><path d="M6.61 6.61A13.526 13.526 0 0 0 2 12s3 7 10 7a9.74 9.74 0 0 0 5.39-1.61"></path><line x1="2" y1="2" x2="22" y2="22"></line></svg>`,
  lock: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg>`,
  unlock: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 9.9-1"></path></svg>`,
  shieldCheck: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path><path d="m9 12 2 2 4-4"></path></svg>`,
  backspace: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 4H8l-7 8 7 8h13a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2z"></path><line x1="18" y1="9" x2="12" y2="15"></line><line x1="12" y1="9" x2="18" y2="15"></line></svg>`,
};

function formatMaskedAccount(accountNo, id, allowHtml = false) {
  if (!accountNo) return '—';
  const strId = String(id || '');
  const isUnmasked = !state.privacyMode || (strId && state.unmaskedIds.has(strId));
  if (isUnmasked) {
    return allowHtml ? `<span class="acct-val unmasked">${escapeHtml(accountNo)}</span>` : accountNo;
  }
  const clean = String(accountNo).trim();
  const last4 = clean.length > 4 ? clean.slice(-4) : clean;
  const maskedText = clean.length > 8 ? `•••• •••• ${last4}` : (clean.length > 4 ? `•••• ${last4}` : `••••`);
  return allowHtml ? `<span class="acct-val masked">${escapeHtml(maskedText)}</span>` : maskedText;
}

function renderAccountWithEye(accountNo, id, extraClass = '') {
  if (!accountNo) return '—';
  const strId = String(id || '');
  const isUnmasked = !state.privacyMode || (strId && state.unmaskedIds.has(strId));
  const displayText = formatMaskedAccount(accountNo, strId, false);
  return `
    <span class="acct-eye-box ${extraClass}" data-stop>
      <span class="mono acct-digits ${isUnmasked ? 'is-unmasked' : 'is-masked'}">${escapeHtml(displayText)}</span>
      <button type="button" class="icon-btn-eye ${isUnmasked ? 'unmasked' : ''}" data-action="toggle-unmask-item" data-id="${escapeHtml(strId)}" title="${isUnmasked ? 'Mask Account Number' : 'Show Full Account Number'}" data-stop>
        ${isUnmasked ? ICONS.eyeOff : ICONS.eye}
      </button>
    </span>
  `;
}

function renderLockScreen() {
  return `
    <div class="lock-screen-wrapper">
      <div class="lock-card">
        <div class="lock-brand">
          <span class="lock-seal">⛁</span>
          <span class="lock-brand-name">Payment</span>
        </div>
        
        <div class="lock-header">
          <h2 class="lock-title">Enter Passcode</h2>
          <p class="lock-sub">Enter your 4-digit security PIN</p>
        </div>
        
        <div class="pin-dots-container ${state.pinError ? 'shake error' : ''} ${state.pinSuccess ? 'success' : ''}">
          <div class="pin-dot ${state.pinInput.length >= 1 ? 'filled' : ''}"></div>
          <div class="pin-dot ${state.pinInput.length >= 2 ? 'filled' : ''}"></div>
          <div class="pin-dot ${state.pinInput.length >= 3 ? 'filled' : ''}"></div>
          <div class="pin-dot ${state.pinInput.length >= 4 ? 'filled' : ''}"></div>
        </div>
        
        <div class="pin-msg-area">
          ${state.pinError ? `<div class="pin-error-msg">${escapeHtml(state.pinError)}</div>` : '<div class="pin-hint-msg">Tap keypad or type digits on keyboard</div>'}
        </div>
        
        <div class="keypad-grid phone-keypad">
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="1">
            <span class="btn-digit">1</span>
            <span class="btn-sub">&nbsp;</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="2">
            <span class="btn-digit">2</span>
            <span class="btn-sub">A B C</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="3">
            <span class="btn-digit">3</span>
            <span class="btn-sub">D E F</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="4">
            <span class="btn-digit">4</span>
            <span class="btn-sub">G H I</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="5">
            <span class="btn-digit">5</span>
            <span class="btn-sub">J K L</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="6">
            <span class="btn-digit">6</span>
            <span class="btn-sub">M N O</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="7">
            <span class="btn-digit">7</span>
            <span class="btn-sub">P Q R S</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="8">
            <span class="btn-digit">8</span>
            <span class="btn-sub">T U V</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="9">
            <span class="btn-digit">9</span>
            <span class="btn-sub">W X Y Z</span>
          </button>
          <button type="button" class="keypad-btn phone-btn keypad-func" data-action="pin-clear" title="Clear">
            <span class="btn-func-txt">C</span>
          </button>
          <button type="button" class="keypad-btn phone-btn" data-action="pin-digit" data-num="0">
            <span class="btn-digit">0</span>
            <span class="btn-sub">+</span>
          </button>
          <button type="button" class="keypad-btn phone-btn keypad-func" data-action="pin-backspace" title="Delete">
            <span class="btn-func-icon">${ICONS.backspace}</span>
          </button>
        </div>

        <div class="lock-foot-actions">
          <button type="button" class="btn-text-ghost" data-action="forgot-pin">Forgot PIN?</button>
        </div>
      </div>
      ${state.modal ? renderModal() : ''}
    </div>
  `;
}

function render() {
  const app = document.getElementById('app');
  
  // Persist current route
  localStorage.setItem(STORE_KEYS.route, state.route);
  
  if (state.isLocked) {
    app.innerHTML = renderLockScreen();
    attachHandlers();
    return;
  }
  
  // Preserve scroll positions of window, modals, tables, and scrollable containers
  const winScrollY = window.scrollY || document.documentElement.scrollTop || 0;
  const winScrollX = window.scrollX || document.documentElement.scrollLeft || 0;
  
  const idScrolls = {};
  app.querySelectorAll('[id]').forEach(el => {
    if (el.id && (el.scrollTop > 0 || el.scrollLeft > 0)) {
      idScrolls[el.id] = { top: el.scrollTop, left: el.scrollLeft };
    }
  });

  const classScrolls = [];
  const scrollSelectors = '.modal, .card-pad, .desktop-only, .amount-sheet, .mobile-cards-list, .mobile-amount-list, .custom-dropdown-options-list, .excel-preview-container';
  app.querySelectorAll(scrollSelectors).forEach((el, index) => {
    if (el.scrollTop > 0 || el.scrollLeft > 0) {
      classScrolls.push({ index, top: el.scrollTop, left: el.scrollLeft });
    }
  });
  
  app.innerHTML = `
    <div class="app-shell">
      <header class="mobile-topbar">
        <div class="mobile-brand">
          <span class="mobile-seal">⛁</span>
          <span class="mobile-title">Payment</span>
        </div>
        <div class="mobile-topbar-actions">
          <button class="topbar-sec-btn ${state.privacyMode ? 'active' : ''}" data-action="toggle-global-privacy" title="${state.privacyMode ? 'Privacy Mode: Account numbers masked' : 'Privacy Mode: Full numbers visible'}">
            ${state.privacyMode ? ICONS.eye : ICONS.eyeOff}
            <span class="topbar-btn-text">${state.privacyMode ? 'Masked' : 'Visible'}</span>
          </button>
          <button class="topbar-sec-btn ${state.pin ? 'pin-active' : ''}" data-action="open-pin-settings" title="${state.pin ? 'PIN Security Active' : 'Set 4-Digit PIN'}">
            ${state.pin ? ICONS.lock : ICONS.key}
          </button>
          <div class="mobile-status-pill">
            <span class="dot"></span>
            <span>${state.parties.length}</span>
          </div>
        </div>
      </header>
      ${renderRail()}
      <div class="main">${renderRoute()}</div>
    </div>
    ${state.modal ? renderModal() : ''}
  `;
  
  // Restore scroll positions immediately
  Object.keys(idScrolls).forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.scrollTop = idScrolls[id].top;
      el.scrollLeft = idScrolls[id].left;
    }
  });

  const newScrollEls = app.querySelectorAll(scrollSelectors);
  classScrolls.forEach(pos => {
    const el = newScrollEls[pos.index];
    if (el) {
      el.scrollTop = pos.top;
      el.scrollLeft = pos.left;
    }
  });

  if (winScrollY > 0 || winScrollX > 0) {
    window.scrollTo(winScrollX, winScrollY);
  }

  attachHandlers();
}

function renderRail() {
  const links = [
    { key: 'dashboard', label: 'Analytics Dashboard', shortLabel: 'Dashboard', icon: 'dashboard' },
    { key: 'directory', label: 'Party Directory', shortLabel: 'Parties', icon: 'directory' },
    { key: 'newrun', label: 'New Payment Run', shortLabel: 'New Run', icon: 'newrun' },
    { key: 'history', label: 'History Logs', shortLabel: 'History', icon: 'history' },
    { key: 'accounts', label: 'My Accounts', shortLabel: 'Accounts', icon: 'bank' },
  ];

  return `
    <nav class="rail">
      <div class="rail-brand">
        <div class="mark"><span class="seal">⛁</span> Payment</div>
        <div class="tagline">Bank Payments</div>
      </div>
      <div class="rail-nav">
        ${links.map(l => `
          <button class="rail-link ${state.route === l.key ? 'active' : ''}" data-route="${l.key}">
            ${ICONS[l.icon]}
            <span class="rail-link-label">${l.label}</span>
            <span class="rail-link-short-label">${l.shortLabel}</span>
          </button>
        `).join('')}
      </div>
      <div class="rail-foot">
        <div class="rail-sec-controls">
          <button class="rail-sec-btn ${state.privacyMode ? 'active' : ''}" data-action="toggle-global-privacy" title="Toggle privacy mask on all account numbers">
            ${state.privacyMode ? ICONS.eye : ICONS.eyeOff}
            <span>${state.privacyMode ? 'Privacy: Masked' : 'Privacy: Visible'}</span>
          </button>
          <button class="rail-sec-btn ${state.pin ? 'active' : ''}" data-action="open-pin-settings" title="PIN Security & Lock Screen">
            ${state.pin ? ICONS.lock : ICONS.key}
            <span>${state.pin ? 'PIN Active' : 'Set PIN'}</span>
          </button>
        </div>
        <div class="rail-foot-text">Secure Storage · Bank Payments</div>
      </div>
    </nav>
  `;
}

function renderRoute() {
  switch (state.route) {
    case 'dashboard': return renderDashboard();
    case 'directory': return renderDirectory();
    case 'newrun': return renderNewRun();
    case 'history': return renderHistory();
    case 'accounts': return renderAccounts();
    default: return renderDashboard();
  }
}

function formatShortINR(num) {
  if (num >= 10000000) return (num / 10000000).toFixed(1) + ' Cr';
  if (num >= 100000) return (num / 100000).toFixed(1) + ' L';
  if (num >= 1000) return (num / 1000).toFixed(1) + ' K';
  return num.toString();
}

function renderDashboard() {
  const history = state.history || [];
  
  // KPI Calculations
  const totalOutflow = history.reduce((sum, h) => sum + (h.total || 0), 0);
  const totalRuns = history.length;
  const totalParties = state.parties.length;
  const totalAccounts = state.myAccounts.length;
  const totalChequesCount = history.reduce((sum, h) => sum + (h.parties ? h.parties.length : 0), 0);

  // 1. Monthly Outflow Grouping (Last 6 months)
  const monthlyData = {};
  history.forEach(h => {
    if (!h.date) return;
    const parts = h.date.split('-');
    if (parts.length < 2) return;
    const yearMonth = `${parts[0]}-${parts[1]}`; // "2026-07"
    monthlyData[yearMonth] = (monthlyData[yearMonth] || 0) + (h.total || 0);
  });

  // Get last 6 months keys chronologically
  const monthKeys = Object.keys(monthlyData).sort().slice(-6);
  if (monthKeys.length === 0) {
    const now = new Date();
    const currentYM = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    monthKeys.push(currentYM);
    monthlyData[currentYM] = 0;
  }

  const rawMax = Math.max(...monthKeys.map(k => monthlyData[k]), 0);
  const maxVal = rawMax > 0 ? rawMax * 1.18 : 1000; // Headroom for label

  const formatMonthName = (ym) => {
    const [y, m] = ym.split('-');
    const date = new Date(Number(y), Number(m) - 1, 1);
    return date.toLocaleString('default', { month: 'short', year: '2-digit' });
  };

  const chartHeight = 175;
  const barWidth = 36;
  const gap = 20;
  const totalBarWidth = barWidth + gap;
  const startX = 30;
  const totalSvgWidth = Math.max(340, startX + monthKeys.length * totalBarWidth + 25);
  const baselineY = chartHeight - 32;

  const svgBarsHtml = monthKeys.map((ym, i) => {
    const val = monthlyData[ym] || 0;
    const barHeight = maxVal > 0 ? (val / maxVal) * (baselineY - 26) : 0;
    const x = startX + i * totalBarWidth;
    const y = baselineY - barHeight;
    const label = formatMonthName(ym);
    
    return `
      <g class="chart-group">
        <rect 
          x="${x}" 
          y="${y}" 
          width="${barWidth}" 
          height="${Math.max(barHeight, 2)}" 
          rx="5" 
          fill="url(#barGradient)" 
          class="chart-bar"
        >
          <title>${label}: ₹ ${formatINR(val)}</title>
        </rect>
        <text 
          x="${x + barWidth / 2}" 
          y="${y - 6}" 
          text-anchor="middle" 
          font-size="10px" 
          font-weight="700" 
          fill="var(--primary-text)"
        >
          ${val > 0 ? '₹' + formatShortINR(val) : '₹0'}
        </text>
        <text 
          x="${x + barWidth / 2}" 
          y="${chartHeight - 12}" 
          text-anchor="middle" 
          font-size="11px" 
          fill="var(--secondary-text)" 
          font-weight="600"
        >
          ${label}
        </text>
      </g>
    `;
  }).join('');

  // 2. Bank Accounts Split
  const bankSplit = {};
  history.forEach(h => {
    if (!h.account) return;
    const key = `${h.account.bankName.toUpperCase()} · ${h.account.accountNo}`;
    bankSplit[key] = (bankSplit[key] || 0) + (h.total || 0);
  });
  const bankSplitSorted = Object.entries(bankSplit)
    .sort((a, b) => b[1] - a[1]);

  // 3. Top Payees
  const payeeSplit = {};
  history.forEach(h => {
    (h.parties || []).forEach(p => {
      const name = (p.name || '').toUpperCase().trim();
      if (!name) return;
      payeeSplit[name] = (payeeSplit[name] || 0) + (p.amount || 0);
    });
  });
  const topPayeesSorted = Object.entries(payeeSplit)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  const maxPayeeVal = topPayeesSorted[0] ? topPayeesSorted[0][1] : 1;

  // 4. Recent Payment Runs (Latest 4)
  const recentRuns = [...history].reverse().slice(0, 4);

  return `
    <div class="page-head">
      <div>
        <h1>Dashboard</h1>
      </div>
      <div class="head-actions desktop-only">
        <button class="btn btn-accent" data-action="goto-newrun">${ICONS.plus} New Payment Run</button>
      </div>
    </div>

    <!-- Dashboard Quick Actions -->
    <div class="dashboard-quick-actions">
      <button type="button" class="dashboard-quick-btn primary" data-action="goto-newrun">
        ${ICONS.plus}
        <span>New Payment Run</span>
      </button>
      <button type="button" class="dashboard-quick-btn" data-action="goto-directory">
        ${ICONS.directory}
        <span>Directory (${totalParties})</span>
      </button>
      <button type="button" class="dashboard-quick-btn" data-action="goto-history">
        ${ICONS.history}
        <span>History (${totalRuns})</span>
      </button>
      <button type="button" class="dashboard-quick-btn" data-action="goto-accounts">
        ${ICONS.bank}
        <span>Accounts (${totalAccounts})</span>
      </button>
      <button type="button" class="dashboard-quick-btn" data-action="open-export-monthly-modal">
        ${ICONS.download}
        <span>Export Month</span>
      </button>
    </div>

    <!-- KPI Grid -->
    <div class="dashboard-kpi-grid">
      <div class="kpi-card">
        <div class="kpi-icon outflow">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><polyline points="19 12 12 19 5 12"></polyline></svg>
        </div>
        <div class="kpi-info">
          <div class="kpi-label-row">
            <span class="kpi-label">Total Disbursed</span>
          </div>
          <h2 class="kpi-value">₹ ${formatINR(totalOutflow)}</h2>
          <span class="kpi-sub">Across all logged runs</span>
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon runs">${ICONS.history}</div>
        <div class="kpi-info">
          <div class="kpi-label-row">
            <span class="kpi-label">Payment Runs</span>
          </div>
          <h2 class="kpi-value">${totalRuns}</h2>
          <span class="kpi-sub">${totalChequesCount} total vouchers</span>
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon parties">${ICONS.directory}</div>
        <div class="kpi-info">
          <div class="kpi-label-row">
            <span class="kpi-label">Active Payees</span>
          </div>
          <h2 class="kpi-value">${totalParties}</h2>
          <span class="kpi-sub">Directory registered</span>
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-icon accounts">${ICONS.bank}</div>
        <div class="kpi-info">
          <div class="kpi-label-row">
            <span class="kpi-label">Bank Accounts</span>
          </div>
          <h2 class="kpi-value">${totalAccounts}</h2>
          <span class="kpi-sub">Connected payer banks</span>
        </div>
      </div>
    </div>

    <!-- Charts Section Grid -->
    <div class="dashboard-charts-grid">
      
      <!-- Outflow Bar Chart Card -->
      <div class="card chart-card">
        <div class="card-head">
          <div style="display:flex; align-items:center; gap:8px;">
            <h3>Monthly Cash Outflow</h3>
            <span class="badge badge-accent" style="font-size:11px;">Last 6 Months</span>
          </div>
        </div>
        <div class="card-pad" style="display:flex; justify-content:center; align-items:center; min-height:200px; padding: 16px 14px; width: 100%; box-sizing: border-box;">
          ${history.length === 0 ? `
            <div class="dashboard-empty-chart" style="text-align:center; padding: 30px 10px;">
              <div style="font-size: 28px; margin-bottom: 8px;">📊</div>
              <div>No transaction logs recorded yet.</div>
              <div style="font-size: 12px; margin-top: 4px; color: var(--secondary-text);">Create your first payment run to view monthly outflow charts.</div>
            </div>
          ` : `
            <svg viewBox="0 0 ${totalSvgWidth} ${chartHeight}" style="width: 100%; max-width: 100%; height: auto; max-height: 200px; display: block;" preserveAspectRatio="xMidYMid meet">
              <defs>
                <linearGradient id="barGradient" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stop-color="#818cf8" />
                  <stop offset="100%" stop-color="#4f46e5" />
                </linearGradient>
              </defs>
              <!-- Gridlines -->
              <line x1="15" y1="26" x2="${totalSvgWidth - 15}" y2="26" stroke="var(--border)" stroke-width="1" stroke-dasharray="3 3" opacity="0.6" />
              <line x1="15" y1="${Math.round(26 + (baselineY - 26) / 2)}" x2="${totalSvgWidth - 15}" y2="${Math.round(26 + (baselineY - 26) / 2)}" stroke="var(--border)" stroke-width="1" stroke-dasharray="3 3" opacity="0.6" />
              <line x1="15" y1="${baselineY}" x2="${totalSvgWidth - 15}" y2="${baselineY}" stroke="var(--border)" stroke-width="1.2" />
              ${svgBarsHtml}
            </svg>
          `}
        </div>
      </div>

      <!-- Bank Accounts Split Card -->
      <div class="card split-card">
        <div class="card-head">
          <div style="display:flex; align-items:center; gap:8px;">
            <h3>Payer Bank Disbursements</h3>
            ${bankSplitSorted.length > 0 ? `<span class="badge" style="font-size:11px;">${bankSplitSorted.length} Banks</span>` : ''}
          </div>
        </div>
        <div class="card-pad scrollable-split-list" style="max-height: 220px; overflow-y: auto;">
          ${bankSplitSorted.length === 0 ? `
            <div class="dashboard-empty-chart" style="text-align:center; padding: 30px 10px;">
              <div style="font-size: 28px; margin-bottom: 8px;">🏦</div>
              <div>No paying bank records found.</div>
            </div>
          ` : bankSplitSorted.map(([acctName, amt]) => {
            const percentage = totalOutflow > 0 ? (amt / totalOutflow) * 100 : 0;
            return `
              <div class="split-row">
                <div class="split-row-meta">
                  <span class="split-name" title="${escapeHtml(acctName)}">${escapeHtml(acctName)}</span>
                  <span class="split-val"><strong>₹ ${formatINR(amt)}</strong> (${percentage.toFixed(0)}%)</span>
                </div>
                <div class="split-progress-container">
                  <div class="split-progress-bar" style="width: ${percentage}%"></div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>

    </div>

    <!-- Bottom Section: Top Payees & Recent Runs -->
    <div class="dashboard-bottom-grid">
      
      <!-- Top Payees Section -->
      <div class="card">
        <div class="card-head">
          <div style="display:flex; align-items:center; gap:8px;">
            <h3>Top 5 Payees by Volume</h3>
            <span class="badge" style="font-size:11px;">Aggregate</span>
          </div>
        </div>
        <div class="card-pad">
          ${topPayeesSorted.length === 0 ? `
            <div class="dashboard-empty-chart" style="padding:24px 0; text-align:center;">
              <div style="font-size: 24px; margin-bottom: 6px;">👥</div>
              <div>No payee payment records found.</div>
            </div>
          ` : `
            <div class="top-payees-list">
              ${topPayeesSorted.map(([payeeName, totalAmt], index) => {
                const pct = maxPayeeVal > 0 ? (totalAmt / maxPayeeVal) * 100 : 0;
                const rankClass = index === 0 ? 'rank-gold' : index === 1 ? 'rank-silver' : index === 2 ? 'rank-bronze' : '';
                return `
                  <div class="top-payee-item">
                    <div class="payee-rank ${rankClass}">${index + 1}</div>
                    <div class="payee-info-col">
                      <span class="payee-name" title="${escapeHtml(payeeName)}">${escapeHtml(payeeName)}</span>
                      <div class="payee-progress-bar-wrap">
                        <div class="payee-progress-bar" style="width: ${pct}%"></div>
                      </div>
                    </div>
                    <div class="payee-amount-col">
                      <strong>₹ ${formatINR(totalAmt)}</strong>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          `}
        </div>
      </div>

      <!-- Recent Payment Runs Preview -->
      <div class="card">
        <div class="card-head">
          <div style="display:flex; align-items:center; justify-content:space-between; width:100%;">
            <div style="display:flex; align-items:center; gap:8px;">
              <h3>Recent Payment Runs</h3>
              ${recentRuns.length > 0 ? `<span class="badge" style="font-size:11px;">Latest ${recentRuns.length}</span>` : ''}
            </div>
            ${recentRuns.length > 0 ? `
              <button type="button" class="btn btn-ghost btn-sm" data-action="goto-history" style="padding: 4px 8px; font-size: 12px; gap: 4px;">
                View All →
              </button>
            ` : ''}
          </div>
        </div>
        <div class="card-pad">
          ${recentRuns.length === 0 ? `
            <div class="dashboard-empty-chart" style="padding:24px 0; text-align:center;">
              <div style="font-size: 24px; margin-bottom: 6px;">📝</div>
              <div>No recent payment runs yet.</div>
              <button class="btn btn-primary btn-sm" data-action="goto-newrun" style="margin-top:10px;">${ICONS.plus} Create First Run</button>
            </div>
          ` : `
            <div class="recent-runs-list">
              ${recentRuns.map(r => `
                <div class="recent-run-item">
                  <div class="recent-run-badge">
                    ${ICONS.newrun}
                  </div>
                  <div class="recent-run-info">
                    <div class="recent-run-title-row">
                      <span class="recent-run-bank">${escapeHtml(r.account?.bankName || 'Payer Bank')}</span>
                      <span class="recent-run-cheque mono">#${escapeHtml(r.chequeNo)}</span>
                    </div>
                    <div class="recent-run-meta">
                      <span>${formatDateDDMMYYYY(r.date)}</span>
                      <span class="bullet">·</span>
                      <span>${(r.parties || []).length} payee${(r.parties || []).length === 1 ? '' : 's'}</span>
                    </div>
                  </div>
                  <div class="recent-run-action">
                    <span class="recent-run-amount">₹ ${formatINR(r.total)}</span>
                    <button class="icon-btn" data-action="view-history" data-id="${r.id}" title="Preview Voucher" style="width:28px; height:28px; padding:0;">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px; height:14px;"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
                    </button>
                  </div>
                </div>
              `).join('')}
            </div>
          `}
        </div>
      </div>

    </div>
  `;
}

function getFilteredParties() {
  const q = state.search.trim().toLowerCase();
  const cat = state.directoryCategory || 'all';
  return state.parties.filter(p => {
    let partyCat = p.category || 'material';
    if (partyCat === 'contractors') partyCat = 'family';
    if (cat !== 'all' && partyCat !== cat) return false;
    return !q || p.name.toLowerCase().includes(q) || p.bankName.toLowerCase().includes(q) || p.accountNo.includes(q);
  });
}

function renderDirectory() {
  const list = getFilteredParties();
  const selectedCount = state.directorySelectedIds.length;
  const allVisibleSelected = list.length > 0 && list.every(p => state.directorySelectedIds.includes(p.id));
  const someVisibleSelected = list.some(p => state.directorySelectedIds.includes(p.id));
  const activeCat = state.directoryCategory || 'all';

  const catCounts = {
    all: state.parties.length,
    material: state.parties.filter(p => (p.category || 'material') === 'material').length,
    employees: state.parties.filter(p => p.category === 'employees').length,
    family: state.parties.filter(p => p.category === 'family' || p.category === 'contractors').length,
    utilities: state.parties.filter(p => p.category === 'utilities').length,
  };

  return `
    <div class="page-head">
      <div>
        <h1>Party Directory</h1>
      </div>
      <div class="head-actions">
        <button class="btn btn-ghost" data-action="open-backup-modal">Backup / Restore</button>
        <button class="btn btn-ghost" data-action="open-import-modal">${ICONS.upload} Import Excel</button>
        <button class="btn btn-primary" data-action="open-party-form">${ICONS.plus} Add Party</button>
      </div>
    </div>

    <!-- Category / Section Filter Tabs -->
    <div class="category-tabs-bar">
      <button type="button" class="cat-tab-btn ${activeCat === 'all' ? 'active' : ''}" data-action="set-directory-category" data-cat="all">
        <span>📋 All Payees</span>
        <span class="cat-tab-count">${catCounts.all}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeCat === 'material' ? 'active' : ''}" data-action="set-directory-category" data-cat="material">
        <span>📦 Material & Vendor Bills</span>
        <span class="cat-tab-count">${catCounts.material}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeCat === 'employees' ? 'active' : ''}" data-action="set-directory-category" data-cat="employees">
        <span>👤 Employees & Staff (Salary)</span>
        <span class="cat-tab-count">${catCounts.employees}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeCat === 'family' ? 'active' : ''}" data-action="set-directory-category" data-cat="family">
        <span>👨‍👩‍👧‍👦 Family</span>
        <span class="cat-tab-count">${catCounts.family}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeCat === 'utilities' ? 'active' : ''}" data-action="set-directory-category" data-cat="utilities">
        <span>🏢 Rent & Utilities</span>
        <span class="cat-tab-count">${catCounts.utilities}</span>
      </button>
    </div>

    <div class="card">
      <div class="card-head">
        <h3>
          ${activeCat === 'all' ? 'All Parties' : getCategoryMeta(activeCat).label} 
          <span class="count">${list.length}</span>
        </h3>
        ${list.length > 0 ? `
          <div class="directory-bulk-actions">
            ${selectedCount > 0 ? `<span class="selection-count">${selectedCount} selected</span>` : ''}
            <button class="btn btn-sm btn-ghost" data-action="toggle-select-all-directory">
              ${allVisibleSelected ? 'Deselect All' : 'Select All'}
            </button>
            ${selectedCount > 0 ? `
              <button class="btn btn-sm btn-danger-ghost" data-action="delete-selected-parties">
                ${ICONS.trash} Delete Selected (${selectedCount})
              </button>
            ` : ''}
          </div>
        ` : ''}
      </div>
      <div class="card-pad" style="padding-bottom:0;">
        <div class="search-bar">
          ${ICONS.search}
          <input type="text" placeholder="Search by name, bank, or account number" value="${escapeHtml(state.search)}" data-bind="search">
        </div>
      </div>
      ${list.length === 0 ? renderDirectoryEmpty(state.search.trim()) : `
        <!-- Desktop Table View -->
        <div class="desktop-only" style="overflow-x:auto;">
          <table>
            <thead>
              <tr>
                <th class="col-check">
                  <div class="check table-check ${allVisibleSelected ? 'checked' : ''} ${someVisibleSelected && !allVisibleSelected ? 'indeterminate' : ''}" data-action="toggle-select-all-directory" title="${allVisibleSelected ? 'Deselect all' : 'Select all'}">
                    ${allVisibleSelected ? ICONS.check : (someVisibleSelected ? '<span class="indeterminate-mark">−</span>' : '')}
                  </div>
                </th>
                <th>Particulars</th>
                <th>Section</th>
                <th>Bank's Name</th>
                <th>Account Number</th>
                <th>Location</th>
                <th>IFSC Code</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${list.map(p => {
                const rowSelected = state.directorySelectedIds.includes(p.id);
                const catMeta = getCategoryMeta(p.category || 'material');
                return `
                <tr class="${rowSelected ? 'row-selected' : ''}">
                  <td class="col-check">
                    <div class="check table-check ${rowSelected ? 'checked' : ''}" data-action="toggle-directory-party" data-id="${p.id}" title="${rowSelected ? 'Deselect' : 'Select'}">
                      ${rowSelected ? ICONS.check : ''}
                    </div>
                  </td>
                  <td><strong>${escapeHtml(p.name)}</strong></td>
                  <td>
                    <span class="cat-badge cat-${catMeta.id}">
                      ${catMeta.icon} ${escapeHtml(catMeta.shortLabel)}
                    </span>
                  </td>
                  <td>${escapeHtml(p.bankName)}</td>
                  <td class="mono">${renderAccountWithEye(p.accountNo, p.id)}</td>
                  <td>${escapeHtml(p.location)}</td>
                  <td class="mono">${escapeHtml(p.ifsc)}</td>
                  <td>
                    <div class="row-actions">
                      <button class="icon-btn" data-action="view-party-history" data-name="${escapeHtml(p.name)}" title="View Payment History">${ICONS.history}</button>
                      <button class="icon-btn" data-action="edit-party" data-id="${p.id}" title="Edit">${ICONS.edit}</button>
                      <button class="icon-btn danger" data-action="delete-party" data-id="${p.id}" title="Delete">${ICONS.trash}</button>
                    </div>
                  </td>
                </tr>
              `;}).join('')}
            </tbody>
          </table>
        </div>

        <!-- Mobile Native Cards View (No horizontal scroll needed!) -->
        <div class="mobile-only mobile-cards-list">
          ${list.map(p => {
            const rowSelected = state.directorySelectedIds.includes(p.id);
            const catMeta = getCategoryMeta(p.category || 'material');
            return `
              <div class="mobile-party-card ${rowSelected ? 'selected' : ''}">
                <div class="mobile-card-header">
                  <div class="check mobile-card-check ${rowSelected ? 'checked' : ''}" data-action="toggle-directory-party" data-id="${p.id}">
                    ${rowSelected ? ICONS.check : ''}
                  </div>
                  <div class="mobile-card-title-wrap">
                    <h4 class="mobile-card-title">${escapeHtml(p.name)}</h4>
                    <span class="cat-badge cat-${catMeta.id}" style="width:fit-content; font-size:11px; padding:2px 8px;">
                      ${catMeta.icon} ${escapeHtml(catMeta.shortLabel)}
                    </span>
                  </div>
                </div>
                <div class="mobile-card-grid">
                  <div class="mobile-card-field">
                    <span class="m-label">Bank</span>
                    <span class="m-val">${escapeHtml(p.bankName)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">Account No</span>
                    <span class="m-val mono">${renderAccountWithEye(p.accountNo, p.id)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">Location</span>
                    <span class="m-val">${escapeHtml(p.location)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">IFSC</span>
                    <span class="m-val mono">${escapeHtml(p.ifsc)}</span>
                  </div>
                </div>
                <div class="mobile-card-actions">
                  <button class="btn btn-sm btn-ghost" data-action="view-party-history" data-name="${escapeHtml(p.name)}">${ICONS.history} History</button>
                  <button class="btn btn-sm btn-ghost" data-action="edit-party" data-id="${p.id}">${ICONS.edit} Edit</button>
                  <button class="btn btn-sm btn-danger-ghost" data-action="delete-party" data-id="${p.id}">${ICONS.trash} Delete</button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `}
    </div>
  `;
}

function renderDirectoryEmpty(hasQuery) {
  if (hasQuery) {
    return renderEmptyState({
      icon: 'search',
      title: 'No Results Found',
      message: 'No parties match your search. Try a different name, bank, or account number.',
      buttons: [{ label: 'Clear Search', action: 'clear-search', variant: 'btn-ghost' }],
    });
  }
  return renderEmptyState({
    icon: 'directory',
    title: 'Directory is Empty',
    message: 'Get started by adding a party manually or importing an Excel list.',
    buttons: [
      { label: 'Import Excel File', action: 'open-import-modal', variant: 'btn-ghost', iconKey: 'upload' },
      { label: 'Add Party', action: 'open-party-form', variant: 'btn-primary', iconKey: 'plus' },
    ],
  });
}

function renderNewRun() {
  const q = state.search.trim().toLowerCase();
  const activeRunCat = state.runCategory || 'all';

  const filteredRunList = state.parties.filter(p => {
    let partyCat = p.category || 'material';
    if (partyCat === 'contractors') partyCat = 'family';
    if (activeRunCat !== 'all' && partyCat !== activeRunCat) return false;
    return !q || p.name.toLowerCase().includes(q) || p.bankName.toLowerCase().includes(q);
  });

  const catCounts = {
    all: state.parties.length,
    material: state.parties.filter(p => (p.category || 'material') === 'material').length,
    employees: state.parties.filter(p => p.category === 'employees').length,
    family: state.parties.filter(p => p.category === 'family' || p.category === 'contractors').length,
    utilities: state.parties.filter(p => p.category === 'utilities').length,
  };

  const selected = state.run.selectedIds
    .map(id => state.parties.find(p => p.id === id))
    .filter(Boolean);

  const total = selected.reduce((sum, p) => sum + (Number(state.run.amounts[p.id]) || 0), 0);
  const account = state.myAccounts.find(a => a.id === state.run.accountId);
  const isEditing = state.run.editingHistoryId !== null;
  const chqInfo = getChequeBookInfo(account);
  const dupRecord = findDuplicateCheque(state.run.chequeNo, state.run.editingHistoryId);

  const allVisibleSelected = filteredRunList.length > 0 && filteredRunList.every(p => state.run.selectedIds.includes(p.id));
  const someVisibleSelected = filteredRunList.some(p => state.run.selectedIds.includes(p.id));

  return `
    <div class="page-head">
      <div>
        <h1>New Payment Run</h1>
      </div>
      <div style="display:flex; gap:12px;">
        ${isEditing ? `<button class="btn btn-ghost" data-action="cancel-editing-run">Cancel Edit</button>` : ''}
        ${selected.length > 0 ? `<button class="btn btn-ghost" data-action="clear-run">Clear Selection</button>` : ''}
      </div>
    </div>

    ${isEditing ? `
      <div style="background: rgba(245, 158, 11, 0.1); border: 1px solid var(--accent); padding: 14px 20px; border-radius: 12px; margin-bottom: 24px; font-size: 13.5px; color: var(--accent); display:flex; align-items:center; gap:10px;">
        <span>⚠️</span>
        <span>You are currently editing an existing payment run (Cheque: <b>${escapeHtml(state.run.chequeNo)}</b>). Saving will update the existing file record instead of creating a new one.</span>
      </div>
    ` : ''}

    <!-- Category / Section Filter Tabs -->
    <div class="category-tabs-bar">
      <button type="button" class="cat-tab-btn ${activeRunCat === 'all' ? 'active' : ''}" data-action="set-run-category" data-cat="all">
        <span>📋 All Payees</span>
        <span class="cat-tab-count">${catCounts.all}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeRunCat === 'material' ? 'active' : ''}" data-action="set-run-category" data-cat="material">
        <span>📦 Material & Vendor Bills</span>
        <span class="cat-tab-count">${catCounts.material}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeRunCat === 'employees' ? 'active' : ''}" data-action="set-run-category" data-cat="employees">
        <span>👤 Employees & Staff (Salary)</span>
        <span class="cat-tab-count">${catCounts.employees}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeRunCat === 'family' ? 'active' : ''}" data-action="set-run-category" data-cat="family">
        <span>👨‍👩‍👧‍👦 Family</span>
        <span class="cat-tab-count">${catCounts.family}</span>
      </button>
      <button type="button" class="cat-tab-btn ${activeRunCat === 'utilities' ? 'active' : ''}" data-action="set-run-category" data-cat="utilities">
        <span>🏢 Rent & Utilities</span>
        <span class="cat-tab-count">${catCounts.utilities}</span>
      </button>
    </div>

    <div class="card">
      <div class="card-head">
        <h3>
          ${activeRunCat === 'all' ? 'All Parties' : getCategoryMeta(activeRunCat).label} 
          <span class="count">${filteredRunList.length}</span>
        </h3>
        ${filteredRunList.length > 0 ? `
          <div class="directory-bulk-actions">
            ${selected.length > 0 ? `<span class="selection-count">${selected.length} selected</span>` : ''}
            <button type="button" class="btn btn-sm btn-ghost" data-action="toggle-select-all-run">
              ${allVisibleSelected ? 'Deselect All' : 'Select All'}
            </button>
            ${selected.length > 0 ? `
              <button type="button" class="btn btn-sm btn-danger-ghost" data-action="clear-run">
                Clear All (${selected.length})
              </button>
            ` : ''}
          </div>
        ` : ''}
      </div>
      <div class="card-pad" style="padding-bottom:0;">
        <div class="search-bar">
          ${ICONS.search}
          <input type="text" placeholder="Search by name, bank, or account number" value="${escapeHtml(state.search)}" data-bind="search">
        </div>
      </div>
      ${filteredRunList.length === 0 ? (state.parties.length === 0 ? renderEmptyState({
        icon: 'empty',
        title: 'No Parties Available',
        message: 'Your party directory is empty. Add parties first before starting a payment run.',
        buttons: [{ label: 'Go to Party Directory', action: 'goto-directory', variant: 'btn-primary' }],
      }) : renderEmptyState({
        icon: 'empty',
        title: 'No Payees Found in this Section',
        message: 'No parties match the selected category or search query.',
      })) : `
        <!-- Desktop Table View (Same as Directory with compact scrollable container) -->
        <div class="desktop-only" id="run-party-table-wrap" style="overflow-x: auto; max-height: 480px; overflow-y: auto; border-top: 1px solid var(--border);">
          <table style="margin: 0;">
            <thead style="position: sticky; top: 0; z-index: 2; background: #ffffff; box-shadow: 0 1px 0 var(--border);">
              <tr>
                <th class="col-check">
                  <div class="check table-check ${allVisibleSelected ? 'checked' : ''} ${someVisibleSelected && !allVisibleSelected ? 'indeterminate' : ''}" data-action="toggle-select-all-run" title="${allVisibleSelected ? 'Deselect all' : 'Select all'}">
                    ${allVisibleSelected ? ICONS.check : (someVisibleSelected ? '<span class="indeterminate-mark">−</span>' : '')}
                  </div>
                </th>
                <th>Particulars</th>
                <th>Section</th>
                <th>Bank's Name</th>
                <th>Account Number</th>
                <th>Location</th>
                <th>IFSC Code</th>
              </tr>
            </thead>
            <tbody>
              ${filteredRunList.map(p => {
                const rowSelected = state.run.selectedIds.includes(p.id);
                const catMeta = getCategoryMeta(p.category || 'material');
                return `
                <tr class="${rowSelected ? 'row-selected' : ''}">
                  <td class="col-check">
                    <div class="check table-check ${rowSelected ? 'checked' : ''}" data-action="toggle-party" data-id="${p.id}" title="${rowSelected ? 'Deselect' : 'Select'}">
                      ${rowSelected ? ICONS.check : ''}
                    </div>
                  </td>
                  <td><strong>${escapeHtml(p.name)}</strong></td>
                  <td>
                    <span class="cat-badge cat-${catMeta.id}">
                      ${catMeta.icon} ${escapeHtml(catMeta.shortLabel)}
                    </span>
                  </td>
                  <td>${escapeHtml(p.bankName)}</td>
                  <td class="mono">${renderAccountWithEye(p.accountNo, p.id)}</td>
                  <td>${escapeHtml(p.location)}</td>
                  <td class="mono">${escapeHtml(p.ifsc)}</td>
                </tr>
              `;}).join('')}
            </tbody>
          </table>
        </div>

        <!-- Mobile Native Cards View -->
        <div class="mobile-only mobile-cards-list" id="run-party-mobile-wrap" style="max-height: 440px; overflow-y: auto; padding: 12px; border-top: 1px solid var(--border);">
          ${filteredRunList.map(p => {
            const rowSelected = state.run.selectedIds.includes(p.id);
            const catMeta = getCategoryMeta(p.category || 'material');
            return `
              <div class="mobile-party-card ${rowSelected ? 'selected' : ''}">
                <div class="mobile-card-header">
                  <div class="check mobile-card-check ${rowSelected ? 'checked' : ''}" data-action="toggle-party" data-id="${p.id}">
                    ${rowSelected ? ICONS.check : ''}
                  </div>
                  <div class="mobile-card-title-wrap">
                    <h4 class="mobile-card-title">${escapeHtml(p.name)}</h4>
                    <span class="cat-badge cat-${catMeta.id}" style="width:fit-content; font-size:11px; padding:2px 8px;">
                      ${catMeta.icon} ${escapeHtml(catMeta.shortLabel)}
                    </span>
                  </div>
                </div>
                <div class="mobile-card-grid">
                  <div class="mobile-card-field">
                    <span class="m-label">Bank</span>
                    <span class="m-val">${escapeHtml(p.bankName)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">Account No</span>
                    <span class="m-val mono">${renderAccountWithEye(p.accountNo, p.id)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">Location</span>
                    <span class="m-val">${escapeHtml(p.location)}</span>
                  </div>
                  <div class="mobile-card-field">
                    <span class="m-label">IFSC</span>
                    <span class="m-val mono">${escapeHtml(p.ifsc)}</span>
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `}
    </div>

    ${selected.length > 0 ? `
      <div class="card amount-entry-card">
        <div class="card-head amount-entry-head">
          <h3>Enter Amounts <span class="count">${selected.length} payee${selected.length === 1 ? '' : 's'}</span></h3>
          <div class="amount-entry-total" id="amount-sheet-total">
            Total: <strong>₹ ${formatINR(total)}</strong>
          </div>
        </div>
        <!-- Desktop Amount Sheet -->
        <div class="desktop-only amount-sheet" id="run-amount-sheet-wrap" style="max-height: 320px; overflow-y: auto;">
          <div class="amount-sheet-head" style="position: sticky; top: 0; z-index: 2; background: #f8fafc; box-shadow: 0 1px 0 var(--border);">
            <span>#</span>
            <span>Payee / Party</span>
            <span>Bank Details</span>
            <span class="right">Amount</span>
            <span></span>
          </div>
          ${selected.map((p, i) => `
            <div class="amount-sheet-row">
              <span class="srno">${i + 1}</span>
              <div class="party-cell">
                <div class="pname">${escapeHtml(p.name)}</div>
              </div>
              <div class="bank-cell">
                <div>${escapeHtml(p.bankName)}</div>
                <div class="acct">${renderAccountWithEye(p.accountNo, p.id)}</div>
              </div>
              <div class="amount-cell">
                <div class="currency-input">
                  <span class="currency-symbol">₹</span>
                  <input
                    type="number"
                    min="0"
                    step="1"
                    class="amount-input"
                    placeholder="0"
                    value="${state.run.amounts[p.id] ?? ''}"
                    data-action="set-amount"
                    data-id="${p.id}"
                    inputmode="numeric"
                  >
                </div>
              </div>
              <button class="icon-btn danger remove-from-run" data-action="remove-from-run" data-id="${p.id}" title="Remove from this run">${ICONS.x}</button>
            </div>
          `).join('')}
          <div class="amount-sheet-foot" style="position: sticky; bottom: 0; z-index: 2; background: #ffffff; box-shadow: 0 -1px 0 var(--border);">
            <span class="foot-label">Grand Total</span>
            <span class="foot-total num" id="amount-sheet-foot-total">₹ ${formatINR(total)}</span>
          </div>
        </div>

        <!-- Mobile Amount Cards View (No horizontal scroll!) -->
        <div class="mobile-only mobile-amount-list" style="max-height: 320px; overflow-y: auto; padding: 10px 12px 14px;">
          ${selected.map((p, i) => `
            <div class="mobile-amount-card">
              <div class="m-amt-head">
                <span class="m-amt-num">${i + 1}</span>
                <div class="m-amt-info">
                  <strong>${escapeHtml(p.name)}</strong>
                  <div class="m-amt-sub">${escapeHtml(p.bankName)} · ${renderAccountWithEye(p.accountNo, p.id)}</div>
                </div>
                <button class="icon-btn danger" data-action="remove-from-run" data-id="${p.id}" title="Remove">${ICONS.x}</button>
              </div>
              <div class="m-amt-input-wrap">
                <label>Amount (₹)</label>
                <div class="currency-input">
                  <span class="currency-symbol">₹</span>
                  <input
                    type="number"
                    min="0"
                    step="1"
                    class="amount-input"
                    placeholder="0"
                    value="${state.run.amounts[p.id] ?? ''}"
                    data-action="set-amount"
                    data-id="${p.id}"
                    inputmode="numeric"
                  >
                </div>
              </div>
            </div>
          `).join('')}
          <div class="amount-sheet-foot" style="border-radius: 10px; margin-top: 6px;">
            <span class="foot-label">Grand Total</span>
            <span class="foot-total num">₹ ${formatINR(total)}</span>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-head"><h3>Cheque & Account Details</h3></div>
        <div class="card-pad">
          <div class="field-grid">
            <div class="field">
              <div class="field-header-row">
                <label>Cheque Number</label>
                ${chqInfo && chqInfo.nextUnused ? `
                  <button type="button" class="btn-link-action" data-action="auto-fill-next-cheque" data-next="${escapeHtml(chqInfo.nextUnused)}">
                    ⚡ Auto-Fill Next: <b>#${escapeHtml(chqInfo.nextUnused)}</b>
                  </button>
                ` : ''}
              </div>
              <input 
                type="text" 
                class="${dupRecord ? 'input-warning-highlight' : ''}" 
                placeholder="${chqInfo && chqInfo.nextUnused ? `e.g. ${chqInfo.nextUnused}` : 'e.g. 291383'}" 
                value="${escapeHtml(state.run.chequeNo)}" 
                data-bind="run.chequeNo"
                autocomplete="off"
              >
              ${dupRecord ? `
                <div class="duplicate-cheque-alert">
                  <div class="dup-icon">⚠️</div>
                  <div class="dup-content">
                    <div class="dup-title">Duplicate Cheque Number Warning!</div>
                    <div class="dup-desc">
                      Cheque <b>#${escapeHtml(state.run.chequeNo)}</b> was already issued on <b>${formatDateDDMMYYYY(dupRecord.date)}</b> 
                      for <b>₹ ${formatINR(dupRecord.total)}</b> (${dupRecord.parties.length} payee${dupRecord.parties.length === 1 ? '' : 's'}).
                    </div>
                  </div>
                </div>
              ` : ''}
              ${chqInfo ? `
                <div class="cheque-series-status ${chqInfo.isOutOfLeaves ? 'danger' : (chqInfo.isNearEnd ? 'warning' : '')}">
                  <span class="status-dot"></span>
                  <span>Book Range: <b>#${escapeHtml(chqInfo.startStr)} — #${escapeHtml(chqInfo.endStr)}</b> (${chqInfo.usedCount}/${chqInfo.totalLeaves} used · <b>${chqInfo.remainingLeaves} leaves left</b>)</span>
                  ${chqInfo.isNearEnd ? `<span class="tag-alert">Ending Soon</span>` : ''}
                  ${chqInfo.isOutOfLeaves ? `<span class="tag-alert danger">Exhausted</span>` : ''}
                </div>
              ` : ''}
            </div>
            <div class="field">
              <label>Date</label>
              <input type="date" value="${state.run.date}" data-bind="run.date">
            </div>
            <div class="field full account-picker-field">
              <label>Payer Account (cheque source)</label>
              ${state.myAccounts.length === 0 ? `
                <div class="account-picker-empty">
                  <div class="glyph">${ICONS.bank}</div>
                  <p>No saved accounts yet.</p>
                  <button type="button" class="btn btn-sm btn-primary" data-action="goto-accounts">Add Account in My Accounts</button>
                </div>
              ` : `
                <div class="account-picker">
                  ${state.myAccounts.map(a => {
                    const isSelected = a.id === state.run.accountId;
                    const aChq = getChequeBookInfo(a);
                    return `
                      <div
                        class="account-picker-item ${isSelected ? 'selected' : ''}"
                        role="button"
                        tabindex="0"
                        data-action="select-payer-account"
                        data-id="${a.id}"
                      >
                        <div class="account-picker-check">${isSelected ? ICONS.check : ''}</div>
                        <div class="account-picker-info">
                          <div class="account-holder">${escapeHtml(a.holderName)}</div>
                          <div class="account-meta">
                            ${escapeHtml(a.bankName)} · ${renderAccountWithEye(a.accountNo, a.id)}
                            ${aChq ? ` · <span style="color:var(--accent); font-weight:600;">Book: #${aChq.startStr}—#${aChq.endStr} (${aChq.remainingLeaves} left)</span>` : ''}
                          </div>
                        </div>
                        <div class="account-picker-icon">${ICONS.bank}</div>
                      </div>
                    `;
                  }).join('')}
                </div>
              `}
            </div>
            <div class="field">
              <label>Amount-in-Words Prefix</label>
              <input type="text" placeholder="e.g. INT " value="${escapeHtml(state.run.prefix)}" data-bind="run.prefix">
              <span class="hint">Prefix prepended to the generated words line. (e.g. "INT ")</span>
            </div>
          </div>
        </div>
      </div>

      <div class="cheque-strip">
        <div class="eyebrow">Cheque Preview (Indian Numbering System)</div>
        <div class="words-line">
          ${total > 0
            ? `<span class="words-text">${escapeHtml(amountToWordsLine(total, state.run.prefix))}</span><span class="total-figure">₹ ${formatINR(total)}</span>`
            : `<span class="placeholder">Enter amounts above to generate preview</span>`}
        </div>
        <div class="meta-row">
          <span>Cheque No: <b>${escapeHtml(state.run.chequeNo || '—')}</b></span>
          <span>Date: <b>${state.run.date || '—'}</b></span>
          <span>Drawn on: <b>${account ? escapeHtml(account.bankName) : '—'}</b></span>
        </div>
      </div>

      <div class="form-actions" style="border-top:none;margin-top:18px;">
        <button class="btn btn-accent" data-action="save-payment-run" ${(!state.run.chequeNo || !state.run.date || !state.run.accountId || total <= 0) ? 'disabled' : ''}>
          ${ICONS.save} ${isEditing ? 'Update Payment Run' : 'Save Payment Run'}
        </button>
      </div>
    ` : ''}
  `;
}

function renderHistory() {
  const q = state.search.trim().toLowerCase();
  const sorted = [...state.history]
    .filter(h => {
      // 1. Search Query Filter
      if (q) {
        const matchCheque = h.chequeNo.toLowerCase().includes(q);
        const matchDate = h.date.toLowerCase().includes(q);
        const matchPartyName = h.parties.some(p => p.name.toLowerCase().includes(q));
        const matchPayerBank = h.account && h.account.bankName.toLowerCase().includes(q);
        if (!matchCheque && !matchDate && !matchPartyName && !matchPayerBank) return false;
      }

      // 2. Period Filter
      if (state.historyPeriod && state.historyPeriod !== 'all') {
        if (!h.date) return false;
        
        const hDate = new Date(h.date);
        hDate.setHours(0,0,0,0);
        const now = new Date();
        now.setHours(0,0,0,0);

        if (state.historyPeriod === 'week') {
          const day = now.getDay();
          const diff = now.getDate() - day + (day === 0 ? -6 : 1);
          const startOfWeek = new Date(now);
          startOfWeek.setDate(diff);
          startOfWeek.setHours(0,0,0,0);
          return hDate >= startOfWeek && hDate <= now;
        }
        if (state.historyPeriod === 'month') {
          const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
          return hDate >= startOfMonth && hDate <= now;
        }
        if (state.historyPeriod === 'last-month') {
          const startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
          const endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
          return hDate >= startOfLastMonth && hDate <= endOfLastMonth;
        }
      }
      return true;
    })
    .sort((a, b) => b.createdAt - a.createdAt);

  return `
    <div class="page-head">
      <div>
        <h1>History Logs</h1>
      </div>
      <div class="head-actions">
        <button class="btn btn-ghost" data-action="open-export-monthly-modal" style="border: 1px solid var(--border); margin-right: 8px;">${ICONS.download} Export Month Report</button>
        <button class="btn btn-primary" data-action="goto-newrun">${ICONS.plus} New Payment Run</button>
      </div>
    </div>
    <div class="card">
      <div class="card-pad" style="padding-bottom:8px;">
        <div style="display: flex; gap: 12px; align-items: center; flex-wrap: wrap;">
          <div class="search-bar" style="flex: 1; min-width: 250px; margin-bottom: 0;">
            ${ICONS.search}
            <input type="text" placeholder="Search by party name, cheque number, or date" value="${escapeHtml(state.search)}" data-bind="search">
          </div>
          <div style="min-width: 160px; height: 44px; display: flex; align-items: center;">
            <select data-bind="historyPeriod" style="height: 100%; width: 100%; border: 1px solid var(--border); border-radius: 10px; font-family: inherit; font-size: 14px; padding: 0 12px; outline: none; cursor: pointer; background: #ffffff; color: var(--primary-text); transition: border-color var(--transition);">
              <option value="all" ${state.historyPeriod === 'all' ? 'selected' : ''}>All Time</option>
              <option value="week" ${state.historyPeriod === 'week' ? 'selected' : ''}>This Week</option>
              <option value="month" ${state.historyPeriod === 'month' ? 'selected' : ''}>This Month</option>
              <option value="last-month" ${state.historyPeriod === 'last-month' ? 'selected' : ''}>Last Month</option>
            </select>
          </div>
        </div>
      </div>
      ${sorted.length === 0 ? renderEmptyState({
        icon: 'history',
        title: q ? 'No Results Found' : 'No History Yet',
        message: q
          ? 'No payment sheets match your search. Try a different cheque number, date, or party name.'
          : 'Once you save a payment run, it will appear here for review and re-download.',
        buttons: q
          ? [{ label: 'Clear Search', action: 'clear-search', variant: 'btn-ghost' }]
          : [{ label: 'Start a Payment Run', action: 'goto-newrun', variant: 'btn-primary', iconKey: 'plus' }],
      }) : `
        <!-- Desktop Table View -->
        <div class="desktop-only" style="overflow-x:auto;">
          <table>
            <thead>
              <tr>
                <th>Generation Date</th>
                <th>Cheque No.</th>
                <th>Payer Account</th>
                <th>Payees</th>
                <th>Total Sum</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${sorted.map(h => {
                const mailSubject = encodeURIComponent(`Combined Cheque Payment Voucher - Cheque No. ${h.chequeNo}`);
                const mailBody = encodeURIComponent(
                  `Dear Sir/Madam,\n\n` +
                  `Please find attached the combined cheque payment split details sheet for the following transaction:\n\n` +
                  `- Payer Bank A/C: ${(h.account ? h.account.holderName : '').toUpperCase()}\n` +
                  `- Bank Name: ${(h.account ? h.account.bankName : '').toUpperCase()}\n` +
                  `- Account Number: ${h.account ? h.account.accountNo : ''}\n` +
                  `- Cheque No: ${h.chequeNo}\n` +
                  `- Date: ${formatDateDDMMYYYY(h.date)}\n` +
                  `- Total Amount: Rs. ${formatINR(h.total)} /-\n` +
                  `- Amount in Words: ${amountToWordsLine(h.total, h.prefix ?? 'INT ')}\n\n` +
                  `[IMPORTANT: Please attach the downloaded file: RTGS-NEFT PAYMENT LIST ${formatDateDDMMYYYY(h.date)}.xlsx to this email before sending]\n\n` +
                  `Best regards,\n` +
                  `${(h.account ? h.account.holderName : '').toUpperCase()}`
                );
                const mailTo = h.account ? h.account.bankEmail : '';
                return `
                <tr class="history-row" data-id="${h.id}" style="cursor:pointer;">
                  <td>${escapeHtml(h.date)}</td>
                  <td class="mono">${escapeHtml(h.chequeNo)}</td>
                  <td>${escapeHtml(h.account ? h.account.bankName : '—')}</td>
                  <td>${h.parties.length} recipient${h.parties.length === 1 ? '' : 's'}</td>
                  <td class="num">₹ ${formatINR(h.total)}</td>
                  <td>
                    <div class="row-actions">
                      <button class="icon-btn" data-action="redownload-history" data-id="${h.id}" title="Download Excel" data-stop>${ICONS.download}</button>
                      <a href="mailto:${escapeHtml(mailTo)}?subject=${mailSubject}&body=${mailBody}" class="icon-btn" title="Draft Email to Bank" data-stop style="display:inline-flex; align-items:center; justify-content:center; text-decoration:none;">
                        ${ICONS.mail}
                      </a>
                      <button class="icon-btn danger" data-action="delete-history" data-id="${h.id}" title="Remove from History" data-stop>${ICONS.trash}</button>
                    </div>
                  </td>
                </tr>
              `;}).join('')}
            </tbody>
          </table>
        </div>

        <!-- Mobile History Cards View (No horizontal scroll!) -->
        <div class="mobile-only mobile-cards-list" style="padding: 10px 12px 14px;">
          ${sorted.map(h => {
            const mailSubject = encodeURIComponent(`Combined Cheque Payment Voucher - Cheque No. ${h.chequeNo}`);
            const mailBody = encodeURIComponent(
              `Dear Sir/Madam,\n\n` +
              `Please find attached the combined cheque payment split details sheet for the following transaction:\n\n` +
              `- Payer Bank A/C: ${(h.account ? h.account.holderName : '').toUpperCase()}\n` +
              `- Bank Name: ${(h.account ? h.account.bankName : '').toUpperCase()}\n` +
              `- Account Number: ${h.account ? h.account.accountNo : ''}\n` +
              `- Cheque No: ${h.chequeNo}\n` +
              `- Date: ${formatDateDDMMYYYY(h.date)}\n` +
              `- Total Amount: Rs. ${formatINR(h.total)} /-\n` +
              `- Amount in Words: ${amountToWordsLine(h.total, h.prefix ?? 'INT ')}\n\n` +
              `[IMPORTANT: Please attach the downloaded file: RTGS-NEFT PAYMENT LIST ${formatDateDDMMYYYY(h.date)}.xlsx to this email before sending]\n\n` +
              `Best regards,\n` +
              `${(h.account ? h.account.holderName : '').toUpperCase()}`
            );
            const mailTo = h.account ? h.account.bankEmail : '';
            return `
              <div class="mobile-history-card history-row" data-id="${h.id}">
                <div class="m-hist-top">
                  <div>
                    <span class="m-hist-date">📅 ${formatDateDDMMYYYY(h.date)}</span>
                    <div class="m-hist-cheque mono">Cheque #${escapeHtml(h.chequeNo)}</div>
                  </div>
                  <div class="m-hist-total num">₹ ${formatINR(h.total)}</div>
                </div>
                <div class="m-hist-details">
                  <div><span class="m-label">Bank:</span> <b>${escapeHtml(h.account ? h.account.bankName : '—')}</b></div>
                  <div><span class="m-label">Payees:</span> <b>${h.parties.length} recipient${h.parties.length === 1 ? '' : 's'}</b></div>
                </div>
                <div class="mobile-card-actions">
                  <button class="btn btn-sm btn-accent" data-action="redownload-history" data-id="${h.id}" data-stop>${ICONS.download} Excel</button>
                  <a href="mailto:${escapeHtml(mailTo)}?subject=${mailSubject}&body=${mailBody}" class="btn btn-sm btn-ghost" data-stop style="text-decoration:none;">${ICONS.mail} Email</a>
                  <button class="btn btn-sm btn-danger-ghost" data-action="delete-history" data-id="${h.id}" data-stop>${ICONS.trash} Delete</button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `}
    </div>
  `;
}

function renderAccounts() {
  return `
    <div class="page-head">
      <div>
        <h1>My Accounts</h1>
      </div>
      <button class="btn btn-primary" data-action="open-account-form">${ICONS.plus} Add Account</button>
    </div>
    <div class="card">
      ${state.myAccounts.length === 0 ? renderEmptyState({
        icon: 'bank',
        title: 'No Accounts Yet',
        message: 'Add your paying accounts to start generating vouchers.',
        buttons: [{ label: 'Add Account', action: 'open-account-form', variant: 'btn-primary', iconKey: 'plus' }],
      }) : `
        <!-- Desktop Table View -->
        <div class="desktop-only" style="overflow-x:auto;">
          <table>
            <thead>
              <tr>
                <th>Account Holder Name</th>
                <th>Bank Name</th>
                <th>Account Number</th>
                <th>Cheque Book Series</th>
                <th>Bank Email</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${state.myAccounts.map(a => {
                const chq = getChequeBookInfo(a);
                const pct = chq && chq.totalLeaves > 0 ? (chq.usedCount / chq.totalLeaves) * 100 : 0;
                return `
                <tr>
                  <td><strong>${escapeHtml(a.holderName)}</strong></td>
                  <td>${escapeHtml(a.bankName)}</td>
                  <td class="mono">${renderAccountWithEye(a.accountNo, a.id)}</td>
                  <td>
                    ${chq ? `
                      <div>
                        <div style="font-weight:600; font-family:'SFMono-Regular',Consolas,monospace; font-size:12.5px; color:var(--primary-text);">
                          #${escapeHtml(chq.startStr)} — #${escapeHtml(chq.endStr)}
                        </div>
                        <div style="font-size:11.5px; color:var(--secondary-text); margin-top:2px;">
                          ${chq.usedCount}/${chq.totalLeaves} used · <b>${chq.remainingLeaves} left</b>
                        </div>
                        <div class="chq-mini-progress">
                          <div class="chq-mini-bar ${chq.isOutOfLeaves ? 'danger' : (chq.isNearEnd ? 'warning' : '')}" style="width:${pct}%;"></div>
                        </div>
                      </div>
                    ` : `
                      <span style="color:var(--secondary-text); font-size:12.5px;">— No series set</span>
                    `}
                  </td>
                  <td>${escapeHtml(a.bankEmail || '—')}</td>
                  <td>
                    <div class="row-actions">
                      <button class="icon-btn" data-action="edit-account" data-id="${a.id}" title="Edit">${ICONS.edit}</button>
                      <button class="icon-btn danger" data-action="delete-account" data-id="${a.id}" title="Delete">${ICONS.trash}</button>
                    </div>
                  </td>
                </tr>
              `;}).join('')}
            </tbody>
          </table>
        </div>

        <!-- Mobile Accounts Cards View (No horizontal scroll!) -->
        <div class="mobile-only mobile-cards-list" style="padding: 10px 12px 14px;">
          ${state.myAccounts.map(a => {
            const chq = getChequeBookInfo(a);
            const pct = chq && chq.totalLeaves > 0 ? (chq.usedCount / chq.totalLeaves) * 100 : 0;
            return `
              <div class="mobile-account-card">
                <div class="m-acct-head">
                  <div>
                    <h4 class="m-acct-name">${escapeHtml(a.holderName)}</h4>
                    <div class="m-acct-bank">${escapeHtml(a.bankName)}</div>
                  </div>
                  <div class="icon-bubble">${ICONS.bank}</div>
                </div>
                <div class="m-acct-body">
                  <div class="m-acct-row">
                    <span class="m-label">Account No</span>
                    <span class="mono">${renderAccountWithEye(a.accountNo, a.id)}</span>
                  </div>
                  ${a.bankEmail ? `
                    <div class="m-acct-row">
                      <span class="m-label">Bank Email</span>
                      <span>${escapeHtml(a.bankEmail)}</span>
                    </div>
                  ` : ''}
                  ${chq ? `
                    <div class="m-acct-chq-box" style="margin-top:4px; padding-top:6px; border-top:1px dashed var(--border);">
                      <div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:4px;">
                        <span>Series: <b>#${escapeHtml(chq.startStr)}—#${escapeHtml(chq.endStr)}</b></span>
                        <span style="color:var(--accent); font-weight:700;">${chq.remainingLeaves} leaves left</span>
                      </div>
                      <div class="chq-mini-progress" style="width:100%;">
                        <div class="chq-mini-bar ${chq.isOutOfLeaves ? 'danger' : (chq.isNearEnd ? 'warning' : '')}" style="width:${pct}%;"></div>
                      </div>
                    </div>
                  ` : ''}
                </div>
                <div class="mobile-card-actions">
                  <button class="btn btn-sm btn-ghost" data-action="edit-account" data-id="${a.id}">${ICONS.edit} Edit</button>
                  <button class="btn btn-sm btn-danger-ghost" data-action="delete-account" data-id="${a.id}">${ICONS.trash} Delete</button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `}
    </div>
  `;
}

function renderDatabase() {
  const cfg = getSupabaseConfig();
  const isConfigured = isSupabaseConfigured();
  const status = state.supabase.status;
  const lastSyncStr = state.supabase.lastSync
    ? new Date(state.supabase.lastSync).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
    : 'Never';

  // Format status banner
  let statusBannerHtml = '';
  if (status === 'connected') {
    statusBannerHtml = `
      <div class="db-status-card connected">
        <div class="db-status-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 6 9 17l-5-5"/></svg>
        </div>
        <div class="db-status-content">
          <h4>Supabase Connected & Active</h4>
          <p>Live synchronized with your Supabase PostgreSQL cloud database. All local modifications are backed up in real time.</p>
        </div>
        <div class="db-status-action">
          <span class="db-pill db-pill-success">Cloud Synced 🟢</span>
        </div>
      </div>
    `;
  } else if (status === 'needs_schema') {
    statusBannerHtml = `
      <div class="db-status-card warning">
        <div class="db-status-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
        </div>
        <div class="db-status-content">
          <h4>Database Tables Missing</h4>
          <p>Connected to your Supabase project, but the required tables have not been created yet. Copy and run the SQL setup script below in your Supabase SQL Editor.</p>
        </div>
        <div class="db-status-action">
          <span class="db-pill db-pill-warning">Action Required ⚠️</span>
        </div>
      </div>
    `;
  } else if (status === 'error') {
    statusBannerHtml = `
      <div class="db-status-card error">
        <div class="db-status-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>
        </div>
        <div class="db-status-content">
          <h4>Connection Failed</h4>
          <p>${escapeHtml(state.supabase.message || 'Unable to communicate with Supabase. Please verify your Project URL and Anon API key.')}</p>
        </div>
        <div class="db-status-action">
          <span class="db-pill db-pill-danger">Offline 🔴</span>
        </div>
      </div>
    `;
  } else if (status === 'checking') {
    statusBannerHtml = `
      <div class="db-status-card checking">
        <div class="loader-spinner-inline" style="width:24px; height:24px;"></div>
        <div class="db-status-content">
          <h4>Connecting to Supabase...</h4>
          <p>Verifying project URL and public Anon API key credentials.</p>
        </div>
      </div>
    `;
  } else {
    statusBannerHtml = `
      <div class="db-status-card info">
        <div class="db-status-icon">${ICONS.database}</div>
        <div class="db-status-content">
          <h4>Local Offline Mode</h4>
          <p>Operating in local browser storage. Enter your Supabase credentials below to enable cloud backup and cross-device sync.</p>
        </div>
        <div class="db-status-action">
          <span class="db-pill db-pill-info">Local Only ⚪</span>
        </div>
      </div>
    `;
  }

  const sqlSchema = getSupabaseSchemaSql();

  return `
    <div class="page-head">
      <div>
        <h1>Cloud Database (Supabase)</h1>
        <p class="sub">Connect your free Supabase PostgreSQL database to persist all payee directories, payer bank accounts, and payment logs in the cloud.</p>
      </div>
      <div class="head-actions">
        <button class="btn btn-ghost" data-action="test-supabase-connection" title="Test Connection">
          ${ICONS.refresh} Test Connection
        </button>
      </div>
    </div>

    ${statusBannerHtml}

    <div class="db-grid">
      
      <!-- Card 1: Supabase Configuration -->
      <div class="card">
        <div class="card-head">
          <div style="display:flex; align-items:center; gap:10px;">
            <div class="icon-bubble">${ICONS.key}</div>
            <h3 style="margin:0;">Supabase Project Credentials</h3>
          </div>
        </div>
        <div class="card-pad">
          <form id="supabase-config-form">
            <div class="field" style="margin-bottom:16px;">
              <label>Supabase Project URL</label>
              <input 
                type="url" 
                name="supabaseUrl" 
                id="supabase-url-input"
                placeholder="https://xyzabcdefghijk.supabase.co" 
                value="${escapeHtml(cfg.url)}" 
                required
                style="font-family:monospace; font-size:13.5px;"
              >
              <span class="hint">Found in Supabase Dashboard → Project Settings → API → Project URL</span>
            </div>

            <div class="field" style="margin-bottom:20px;">
              <label>Anon Public API Key</label>
              <input 
                type="password" 
                name="supabaseAnonKey" 
                id="supabase-anon-key-input"
                placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..." 
                value="${escapeHtml(cfg.anonKey)}" 
                required
                style="font-family:monospace; font-size:13.5px;"
              >
              <span class="hint">Found in Supabase Dashboard → Project Settings → API → Project API Keys (anon / public)</span>
            </div>

            <div class="form-actions" style="display:flex; gap:10px; align-items:center;">
              <button type="submit" class="btn btn-primary">
                ${ICONS.save} Save & Connect
              </button>
              ${isConfigured ? `
                <button type="button" class="btn btn-ghost" data-action="clear-supabase-config" style="color:var(--danger);">
                  Disconnect
                </button>
              ` : ''}
            </div>
          </form>
        </div>
      </div>

      <!-- Card 2: Two-way Cloud Synchronization -->
      <div class="card">
        <div class="card-head">
          <div style="display:flex; align-items:center; gap:10px;">
            <div class="icon-bubble">${ICONS.cloud}</div>
            <h3 style="margin:0;">Data Synchronization</h3>
          </div>
          <span class="hint" style="margin:0;">Last Synced: <b>${lastSyncStr}</b></span>
        </div>
        <div class="card-pad">
          <div class="db-stats-summary">
            <div class="db-stat-box">
              <span class="stat-num">${state.parties.length}</span>
              <span class="stat-lbl">Directory Payees</span>
            </div>
            <div class="db-stat-box">
              <span class="stat-num">${state.myAccounts.length}</span>
              <span class="stat-lbl">Bank Accounts</span>
            </div>
            <div class="db-stat-box">
              <span class="stat-num">${state.history.length}</span>
              <span class="stat-lbl">History Logs</span>
            </div>
          </div>

          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:16px; margin-top:20px;">
            <div class="db-sync-card">
              <div class="sync-card-icon">${ICONS.uploadCloud}</div>
              <div class="sync-card-body">
                <h4>Push Local to Cloud</h4>
                <p>Upload all existing browser records (payees, accounts, vouchers) to your Supabase database.</p>
                <button 
                  class="btn btn-primary btn-sm" 
                  data-action="sync-push-to-cloud" 
                  ${!isConfigured ? 'disabled title="Please save your Supabase credentials first"' : ''}
                >
                  🚀 Upload to Cloud
                </button>
              </div>
            </div>

            <div class="db-sync-card">
              <div class="sync-card-icon">${ICONS.downloadCloud}</div>
              <div class="sync-card-body">
                <h4>Pull from Cloud</h4>
                <p>Download the latest cloud database tables and refresh this browser's local cache.</p>
                <button 
                  class="btn btn-ghost btn-sm" 
                  data-action="sync-pull-from-cloud" 
                  ${!isConfigured ? 'disabled title="Please save your Supabase credentials first"' : ''}
                >
                  📥 Download from Cloud
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

    </div>

    <!-- Card 3: 1-Click Database Setup SQL Schema -->
    <div class="card" style="margin-top:24px;">
      <div class="card-head" style="display:flex; justify-content:space-between; align-items:center;">
        <div style="display:flex; align-items:center; gap:10px;">
          <div class="icon-bubble">${ICONS.database}</div>
          <div>
            <h3 style="margin:0;">Database Setup (SQL Script)</h3>
            <span class="hint" style="margin-top:2px;">One-click setup for your Supabase SQL Editor</span>
          </div>
        </div>
        <button class="btn btn-primary" data-action="copy-sql-schema">
          ${ICONS.copy} Copy SQL Setup Script
        </button>
      </div>
      <div class="card-pad">
        <div class="db-setup-steps">
          <div class="setup-step">
            <span class="step-num">1</span>
            <span class="step-text">Open your project at <a href="https://supabase.com/dashboard" target="_blank" rel="noopener noreferrer">supabase.com</a></span>
          </div>
          <div class="setup-step">
            <span class="step-num">2</span>
            <span class="step-text">Go to <b>SQL Editor</b> on the left sidebar & click <b>New Query</b></span>
          </div>
          <div class="setup-step">
            <span class="step-num">3</span>
            <span class="step-text">Click <b>Copy SQL Setup Script</b> above and paste it into the editor</span>
          </div>
          <div class="setup-step">
            <span class="step-num">4</span>
            <span class="step-text">Click <b>RUN</b> in Supabase, then enter your API credentials above!</span>
          </div>
        </div>

        <div class="sql-code-container">
          <div class="sql-code-header">
            <span>PostgreSQL Schema DDL</span>
            <button class="btn btn-sm btn-ghost" data-action="copy-sql-schema" style="height:28px; padding:0 10px; font-size:12px;">
              ${ICONS.copy} Copy
            </button>
          </div>
          <pre class="sql-code-body"><code>${escapeHtml(sqlSchema)}</code></pre>
        </div>
      </div>
    </div>
  `;
}

function renderModal() {
  const m = state.modal;
  if (m.type === 'party-form') return renderPartyFormModal(m.payload);
  if (m.type === 'account-form') return renderAccountFormModal(m.payload);
  if (m.type === 'pin-settings') return renderPinSettingsModal();
  if (m.type === 'confirm-reset-pin') return renderConfirmModal('Reset Security PIN?', 'Are you sure you want to reset and disable your 4-digit PIN? Your bank payment data will remain completely safe and intact.', 'reset-pin-confirmed');
  if (m.type === 'confirm-duplicate-cheque') return renderDuplicateChequeConfirmModal(m.payload);
  if (m.type === 'confirm-delete-party') return renderConfirmModal('Delete Payee?', `Are you sure you want to remove "${escapeHtml(m.payload.name)}" from your directory?`, 'delete-party-confirmed', m.payload.id);
  if (m.type === 'confirm-delete-parties-bulk') {
    const count = m.payload.ids.length;
    const preview = m.payload.names.slice(0, 3).map(n => escapeHtml(n)).join(', ');
    const more = count > 3 ? ` and ${count - 3} more` : '';
    return renderConfirmModal(
      `Delete ${count} Part${count === 1 ? 'y' : 'ies'}?`,
      `Are you sure you want to remove <b>${count}</b> selected payee${count === 1 ? '' : 's'} from your directory (${preview}${more})? This cannot be undone.`,
      'delete-selected-parties-confirmed'
    );
  }
  if (m.type === 'confirm-delete-account') return renderConfirmModal('Delete Account?', `Are you sure you want to remove "${escapeHtml(m.payload.holderName)}" from My Accounts?`, 'delete-account-confirmed', m.payload.id);
  if (m.type === 'confirm-delete-history') return renderConfirmModal('Delete Log Entry?', `This will permanently remove the payment run log dated ${escapeHtml(m.payload.date)} from your history.`, 'delete-history-confirmed', m.payload.id);
  if (m.type === 'view-history') return renderViewHistoryModal(m.payload);
  if (m.type === 'import-modal') return renderImportModal();
  if (m.type === 'backup-modal') return renderBackupModal();
  if (m.type === 'run-success') return renderRunSuccessModal(m.payload);
  if (m.type === 'party-history') return renderPartyHistoryModal(m.payload);
  if (m.type === 'export-monthly-modal') return renderMonthlyExportModal();
  return '';
}

function renderPinSettingsModal() {
  const hasPin = Boolean(state.pin);
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:440px;" data-stop>
        <div class="modal-head">
          <div style="display:flex; align-items:center; gap:10px;">
            <div style="background:var(--accent-soft); color:var(--accent); width:36px; height:36px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
              ${ICONS.lock}
            </div>
            <div>
              <h3 style="margin:0; font-size:17px;">${hasPin ? 'PIN Security Settings' : 'Set 4-Digit PIN'}</h3>
              <p style="margin:0; font-size:12px; color:var(--secondary-text);">Protect app & bank details from shoulder surfing</p>
            </div>
          </div>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body" style="padding:20px;">
          ${hasPin ? `
            <div style="display:flex; flex-direction:column; gap:16px;">
              <div style="display:flex; align-items:center; justify-content:space-between; padding:12px 16px; background:var(--surface-hover); border-radius:12px; border:1px solid var(--border);">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:20px;">🛡️</span>
                  <div>
                    <div style="font-weight:600; font-size:14px; color:var(--primary-text);">4-Digit PIN Lock Active</div>
                    <div style="font-size:12px; color:var(--secondary-text);">App locks automatically on open/refresh</div>
                  </div>
                </div>
                <button type="button" class="btn btn-sm btn-accent" data-action="lock-app" style="display:flex; align-items:center; gap:6px;">
                  ${ICONS.lock} Lock Now
                </button>
              </div>

              <!-- Change PIN Section -->
              <div class="card" style="padding:16px; border:1px solid var(--border); border-radius:12px; background:var(--surface);">
                <h4 style="margin:0 0 12px; font-size:14px; font-weight:600; color:var(--primary-text);">Change 4-Digit PIN</h4>
                <form id="change-pin-form" style="display:flex; flex-direction:column; gap:12px;">
                  <div class="field" style="margin:0;">
                    <label style="font-size:12px;">Current PIN</label>
                    <input type="password" name="currentPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:6px; font-size:16px; text-align:center;">
                  </div>
                  <div class="field" style="margin:0;">
                    <label style="font-size:12px;">New 4-Digit PIN</label>
                    <input type="password" name="newPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:6px; font-size:16px; text-align:center;">
                  </div>
                  <div class="field" style="margin:0;">
                    <label style="font-size:12px;">Confirm New PIN</label>
                    <input type="password" name="confirmPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:6px; font-size:16px; text-align:center;">
                  </div>
                  <button type="submit" class="btn btn-primary" style="margin-top:4px;">Update PIN</button>
                </form>
              </div>

              <!-- Disable PIN Section -->
              <div class="card" style="padding:16px; border:1px solid rgba(239,68,68,0.2); border-radius:12px; background:rgba(239,68,68,0.02);">
                <h4 style="margin:0 0 6px; font-size:13px; font-weight:600; color:var(--danger);">Disable PIN Protection</h4>
                <p style="font-size:12px; color:var(--secondary-text); margin:0 0 10px;">Remove lock screen when opening the app.</p>
                <form id="disable-pin-form" style="display:flex; gap:8px; align-items:flex-end;">
                  <div class="field" style="margin:0; flex:1;">
                    <label style="font-size:11px;">Current PIN</label>
                    <input type="password" name="currentPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:4px; font-size:14px; text-align:center;">
                  </div>
                  <button type="submit" class="btn btn-danger-ghost" style="height:42px;">Remove PIN</button>
                </form>
              </div>
            </div>
          ` : `
            <form id="setup-pin-form" style="display:flex; flex-direction:column; gap:16px;">
              <div style="text-align:center; padding:10px 0;">
                <div style="font-size:36px; margin-bottom:8px;">🔒</div>
                <p style="font-size:13px; color:var(--secondary-text); margin:0 auto; max-width:320px;">
                  Set a 4-digit PIN to lock your bank payments application from unauthorized eyes.
                </p>
              </div>
              <div class="field" style="margin:0;">
                <label style="font-size:13px; font-weight:600;">Create 4-Digit PIN</label>
                <input type="password" name="newPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:8px; font-size:20px; text-align:center; font-weight:bold;">
              </div>
              <div class="field" style="margin:0;">
                <label style="font-size:13px; font-weight:600;">Confirm 4-Digit PIN</label>
                <input type="password" name="confirmPin" maxlength="4" pattern="[0-9]{4}" inputmode="numeric" required placeholder="••••" style="letter-spacing:8px; font-size:20px; text-align:center; font-weight:bold;">
              </div>
              <div class="form-actions" style="margin-top:8px; padding-top:12px; border-top:1px solid var(--border);">
                <button type="button" class="btn btn-ghost" data-action="close-modal">Cancel</button>
                <button type="submit" class="btn btn-accent">${ICONS.lock} Set & Enable PIN</button>
              </div>
            </form>
          `}
        </div>
      </div>
    </div>
  `;
}

function renderDuplicateChequeConfirmModal(payload) {
  const { record, dupRecord } = payload;
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:480px;" data-stop>
        <div class="modal-head">
          <h3 style="display:flex; align-items:center; gap:8px; color:var(--accent);">
            <span>⚠️</span> Duplicate Cheque Warning
          </h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <p style="font-size:14.5px; line-height:1.5; margin:0 0 16px; color:var(--primary-text);">
            Cheque number <b>#${escapeHtml(record.chequeNo)}</b> has already been issued in past records.
          </p>
          <div style="background:rgba(245,158,11,0.08); border:1px solid var(--accent); border-radius:10px; padding:14px 16px; font-size:13px; line-height:1.5; color:var(--primary-text); margin-bottom:20px;">
            <div><b>Previous Issuance Details:</b></div>
            <div style="color:var(--secondary-text); margin-top:4px;">
              • Date: <b>${formatDateDDMMYYYY(dupRecord.date)}</b><br>
              • Account: <b>${escapeHtml(dupRecord.account?.bankName || '—')}</b><br>
              • Total Amount: <b>₹ ${formatINR(dupRecord.total)}</b> (${dupRecord.parties.length} payee${dupRecord.parties.length === 1 ? '' : 's'})
            </div>
          </div>
          <p style="font-size:13px; color:var(--secondary-text); margin:0 0 20px;">
            Are you sure you want to save another payment run with this same cheque number?
          </p>
          <div class="form-actions" style="border-top:none; padding-top:0; margin-top:0;">
            <button class="btn btn-ghost" data-action="close-modal">Change Cheque No</button>
            <button class="btn btn-accent" data-action="confirm-duplicate-cheque-save">Save Anyway</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderAccountFormModal(payload) {
  const editing = !!payload;
  const a = payload || { holderName: '', bankName: '', accountNo: '', bankEmail: '', chequeBookStart: '', chequeBookEnd: '' };
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" data-stop>
        <div class="modal-head">
          <h3>${editing ? 'Edit Paying Account' : 'Add Paying Account'}</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <form id="account-form">
            <div class="field-grid">
              <div class="field full">
                <label>Name on Bank Account</label>
                <input type="text" name="holderName" required placeholder="e.g. Tribhovan Raghavji Prajapati" value="${escapeHtml(a.holderName)}">
              </div>
              <div class="field">
                <label>Name of Bank</label>
                <input type="text" name="bankName" required placeholder="e.g. Union Bank of India" value="${escapeHtml(a.bankName)}">
              </div>
              <div class="field">
                <label>Bank Account Number</label>
                <input type="text" name="accountNo" required placeholder="e.g. SOD # 107021048004036" value="${escapeHtml(a.accountNo)}">
              </div>
              <div class="field full">
                <label>Bank Email (for sending voucher files)</label>
                <input type="email" name="bankEmail" placeholder="e.g. branchmanager@unionbank.com" value="${escapeHtml(a.bankEmail || '')}">
              </div>

              <div class="field full" style="margin-top:6px; padding-top:14px; border-top:1px dashed var(--border);">
                <div style="font-size:13.5px; font-weight:700; color:var(--primary-text); margin-bottom:4px;">
                  🔢 Cheque Book Series Tracking (Optional)
                </div>
                <div style="font-size:12.5px; color:var(--secondary-text); margin-bottom:12px; line-height:1.4;">
                  Enter the start and end cheque leaf numbers from your physical chequebook to enable automatic sequential auto-fill and duplicate number warnings.
                </div>
              </div>
              <div class="field">
                <label>Cheque Book Start No.</label>
                <input type="text" name="chequeBookStart" placeholder="e.g. 291381" value="${escapeHtml(a.chequeBookStart || '')}">
              </div>
              <div class="field">
                <label>Cheque Book End No.</label>
                <input type="text" name="chequeBookEnd" placeholder="e.g. 291400" value="${escapeHtml(a.chequeBookEnd || '')}">
              </div>
            </div>
            <div class="form-actions">
              <button type="button" class="btn btn-ghost" data-action="close-modal">Cancel</button>
              <button type="submit" class="btn btn-primary">${editing ? 'Save Changes' : 'Add Account'}</button>
            </div>
            ${editing ? `<input type="hidden" name="id" value="${a.id}">` : ''}
          </form>
        </div>
      </div>
    </div>
  `;
}

function renderMonthlyExportModal() {
  const months = [
    { value: 1, label: 'January' },
    { value: 2, label: 'February' },
    { value: 3, label: 'March' },
    { value: 4, label: 'April' },
    { value: 5, label: 'May' },
    { value: 6, label: 'June' },
    { value: 7, label: 'July' },
    { value: 8, label: 'August' },
    { value: 9, label: 'September' },
    { value: 10, label: 'October' },
    { value: 11, label: 'November' },
    { value: 12, label: 'December' }
  ];

  const historyYears = state.history.map(h => {
    if (!h.date) return null;
    const y = new Date(h.date).getFullYear();
    return isNaN(y) ? null : y;
  }).filter(Boolean);
  
  const uniqueYears = Array.from(new Set([new Date().getFullYear(), ...historyYears])).sort((a,b) => b - a);

  const mOpts = months.map(m => `<option value="${m.value}" ${state.monthlyExport.month === m.value ? 'selected' : ''}>${escapeHtml(m.label)}</option>`).join('');
  const yOpts = uniqueYears.map(y => `<option value="${y}" ${state.monthlyExport.year === y ? 'selected' : ''}>${y}</option>`).join('');

  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:440px;" data-stop>
        <div class="modal-head">
          <h3>Export Month's Payments</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <p style="margin:0 0 20px; color:var(--slate); font-size:14px;">Select the month and year you wish to download the combined Excel sheet for.</p>
          <div class="field-grid" style="grid-template-columns: 1fr 1fr; gap: 16px;">
            <div class="field">
              <label>Month</label>
              <select data-bind="monthlyExport.month" style="width: 100%; height: 44px; border: 1px solid var(--border); border-radius: 10px; font-family: inherit; font-size: 14px; padding: 0 12px; outline: none; background: #ffffff; color: var(--primary-text);">
                ${mOpts}
              </select>
            </div>
            <div class="field">
              <label>Year</label>
              <select data-bind="monthlyExport.year" style="width: 100%; height: 44px; border: 1px solid var(--border); border-radius: 10px; font-family: inherit; font-size: 14px; padding: 0 12px; outline: none; background: #ffffff; color: var(--primary-text);">
                ${yOpts}
              </select>
            </div>
          </div>
          <div class="form-actions" style="margin-top:24px; padding-top:20px;">
            <button class="btn btn-ghost" data-action="close-modal">Cancel</button>
            <button class="btn btn-primary" data-action="confirm-export-monthly">Export Excel</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderConfirmModal(title, body, action, id) {
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:420px;" data-stop>
        <div class="modal-head">
          <h3>${escapeHtml(title)}</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <p style="margin:0 0 24px; color:var(--slate); font-size:14px;">${body}</p>
          <div class="form-actions" style="margin-top:0; padding-top:0; border-top:none;">
            <button class="btn btn-ghost" data-action="close-modal">Cancel</button>
            <button class="btn btn-danger" data-action="${action}"${id != null ? ` data-id="${id}"` : ''}>Delete</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderPartyFormModal(payload) {
  const editing = !!payload;
  const p = payload || { name: '', bankName: '', accountNo: '', location: '', ifsc: '', category: 'material' };
  const currentCat = p.category || 'material';

  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:560px;" data-stop>
        <div class="modal-head">
          <h3>${editing ? 'Edit Payee Party' : 'Add New Payee'}</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <form id="party-form">
            <div class="field-grid">
              <div class="field full">
                <label>Select List / Section</label>
                <div class="category-selector-grid">
                  ${PARTY_CATEGORIES.map(c => `
                    <label class="category-radio-card ${currentCat === c.id ? 'selected' : ''}" data-action="select-form-category" data-cat="${c.id}">
                      <input type="radio" name="category" value="${c.id}" ${currentCat === c.id ? 'checked' : ''} style="display:none;">
                      <span class="cat-icon">${c.icon}</span>
                      <span class="cat-label">${c.label}</span>
                    </label>
                  `).join('')}
                </div>
              </div>
              <div class="field full">
                <label>Particulars (Payee / Employee / Vendor Name)</label>
                <input type="text" name="name" required placeholder="e.g. Shree Mahavir Timbers or Ramesh Patel" value="${escapeHtml(p.name)}">
              </div>
              <div class="field">
                <label>Bank Name</label>
                <input type="text" name="bankName" required placeholder="e.g. Bank of Maharashtra" value="${escapeHtml(p.bankName)}">
              </div>
              <div class="field">
                <div class="field-header-row">
                  <label>Bank Account Number</label>
                  <span class="val-badge" id="party-form-acct-val"></span>
                </div>
                <input type="text" name="accountNo" required placeholder="e.g. 2053402589" value="${escapeHtml(p.accountNo)}" autocomplete="off">
              </div>
              <div class="field">
                <label>Bank Location (branch)</label>
                <input type="text" name="location" required placeholder="e.g. Anand" value="${escapeHtml(p.location)}">
              </div>
              <div class="field">
                <div class="field-header-row">
                  <label>IFSC Code</label>
                  <span class="val-badge" id="party-form-ifsc-val"></span>
                </div>
                <input type="text" name="ifsc" required placeholder="e.g. MAHB0006456" value="${escapeHtml(p.ifsc)}" autocomplete="off" style="text-transform: uppercase;">
              </div>
            </div>
            <div class="form-actions">
              <button type="button" class="btn btn-ghost" data-action="close-modal">Cancel</button>
              <button type="submit" class="btn btn-primary">${editing ? 'Save Changes' : 'Add Party'}</button>
            </div>
            ${editing ? `<input type="hidden" name="id" value="${p.id}">` : ''}
          </form>
        </div>
      </div>
    </div>
  `;
}



function renderPartyHistoryModal(partyName) {
  const payments = [];
  state.history.forEach(h => {
    h.parties.forEach(p => {
      if (p.name.toLowerCase() === partyName.toLowerCase()) {
        payments.push({
          date: h.date,
          chequeNo: h.chequeNo,
          accountName: h.account ? h.account.bankName : '—',
          accountNo: h.account ? h.account.accountNo : '',
          amount: p.amount
        });
      }
    });
  });
  
  payments.sort((a, b) => new Date(b.date) - new Date(a.date));
  const totalPaid = payments.reduce((sum, p) => sum + p.amount, 0);

  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:720px;" data-stop>
        <div class="modal-head">
          <h3>Payment History — ${escapeHtml(partyName.toUpperCase())}</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          ${payments.length === 0 ? renderEmptyState({
            icon: 'history',
            title: 'No Payments Yet',
            message: 'No payments recorded for this party in History Logs.',
          }) : `
            <div style="margin-bottom: 16px; font-size:14px; color:var(--secondary-text);">
              Found <b>${payments.length}</b> payment transactions totaling <b style="color:var(--accent);">₹ ${formatINR(totalPaid)}</b>.
            </div>
            <div style="overflow-x:auto; border:1px solid var(--border); border-radius:8px; margin-bottom: 20px;">
              <table style="font-size:13.5px; width:100%; border-collapse:collapse;">
                <thead>
                  <tr style="background: rgba(0,0,0,0.2);">
                    <th style="padding:10px 16px; border-bottom:1px solid var(--border); font-size:11px; text-transform:uppercase; color:var(--secondary-text);">Voucher Date</th>
                    <th style="padding:10px 16px; border-bottom:1px solid var(--border); font-size:11px; text-transform:uppercase; color:var(--secondary-text);">Cheque No.</th>
                    <th style="padding:10px 16px; border-bottom:1px solid var(--border); font-size:11px; text-transform:uppercase; color:var(--secondary-text);">Payer Bank Account</th>
                    <th style="padding:10px 16px; border-bottom:1px solid var(--border); font-size:11px; text-transform:uppercase; color:var(--secondary-text); text-align:right;">Amount Paid</th>
                  </tr>
                </thead>
                <tbody>
                  ${payments.map(p => `
                    <tr>
                      <td style="padding:12px 16px; border-bottom:1px solid var(--border);">${escapeHtml(formatDateDDMMYYYY(p.date))}</td>
                      <td style="padding:12px 16px; border-bottom:1px solid var(--border);" class="mono">${escapeHtml(p.chequeNo)}</td>
                      <td style="padding:12px 16px; border-bottom:1px solid var(--border);">${escapeHtml(p.accountName)} <span style="font-size:11px; color:var(--secondary-text);">(${renderAccountWithEye(p.accountNo, 'ph_' + p.chequeNo)})</span></td>
                      <td style="padding:12px 16px; border-bottom:1px solid var(--border); text-align:right;" class="num">₹ ${formatINR(p.amount)}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          `}
          <div class="form-actions" style="margin-top: 20px; padding-top: 16px; border-top: 1px solid var(--border);">
            <button class="btn btn-primary" data-action="close-modal">Close</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function resetQuickAdd() {
  state.quickAdd = {
    selectedIds: [],
    partyId: '',
    mode: 'directory',
    dropdownOpen: false,
    dropdownSearch: '',
    amounts: {},
    defaultAmount: ''
  };
}

function showViewHistoryModal(h, resetForm = true) {
  if (resetForm) resetQuickAdd();
  openModal({ type: 'view-history', payload: h }, false);
}

function renderQuickAddSection(h) {
  const qa = state.quickAdd;
  if (!qa.selectedIds) qa.selectedIds = qa.partyId ? [qa.partyId] : [];
  if (!qa.amounts) qa.amounts = {};

  const isDirectory = qa.mode !== 'manual';
  const isDropdownOpen = !!qa.dropdownOpen;

  const selectedParties = qa.selectedIds.map(pid => state.parties.find(p => p.id === pid)).filter(Boolean);
  const dropdownSearchVal = qa.dropdownSearch || '';
  const dsQ = dropdownSearchVal.trim().toLowerCase();

  const availableParties = state.parties.filter(p => {
    const alreadyIn = h.parties.some(hp => (hp.accountNo || '').trim().toLowerCase() === (p.accountNo || '').trim().toLowerCase());
    return !alreadyIn;
  });

  const filteredParties = availableParties.filter(p => {
    if (!dsQ) return true;
    return (
      p.name.toLowerCase().includes(dsQ) ||
      p.bankName.toLowerCase().includes(dsQ) ||
      p.accountNo.includes(dsQ) ||
      (p.location && p.location.toLowerCase().includes(dsQ))
    );
  });

  const allFilteredSelected = filteredParties.length > 0 && filteredParties.every(p => qa.selectedIds.includes(p.id));

  const customDropdownOptionsHtml = filteredParties.length === 0
    ? `<div class="custom-dropdown-empty">No matching parties found</div>`
    : filteredParties.map(p => {
        const isSelected = qa.selectedIds.includes(p.id);
        const catMeta = getCategoryMeta(p.category || 'material');
        return `
          <div class="custom-dropdown-option ${isSelected ? 'selected' : ''}" data-action="toggle-quick-add-party-select" data-id="${p.id}" data-search="${escapeHtml((p.name + ' ' + p.bankName + ' ' + p.accountNo + ' ' + (p.location || '')).toLowerCase())}">
            <div class="opt-checkbox">
              ${isSelected ? `<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="3" fill="none"><path d="M20 6 9 17l-5-5"/></svg>` : ''}
            </div>
            <div class="option-main" style="flex:1; min-width:0;">
              <span class="option-name">${escapeHtml(p.name)}</span>
              <span class="option-bank">${escapeHtml(p.bankName)} · ${escapeHtml(p.location)} · <span class="mono">${renderAccountWithEye(p.accountNo, p.id)}</span></span>
            </div>
            <span class="cat-badge cat-${catMeta.id}" style="font-size:10px; padding:2px 7px; margin-left:auto; flex-shrink:0;">${catMeta.icon} ${escapeHtml(catMeta.shortLabel)}</span>
          </div>
        `;
      }).join('');

  let triggerText = '-- Choose Parties from Directory --';
  if (selectedParties.length === 1) {
    triggerText = `${escapeHtml(selectedParties[0].name)} (${escapeHtml(selectedParties[0].bankName)})`;
  } else if (selectedParties.length > 1) {
    triggerText = `✓ ${selectedParties.length} Parties Selected`;
  }

  const dropdownHtml = `
    <div class="quick-add-custom-select-container ${isDropdownOpen ? 'open' : ''}">
      <button type="button" class="quick-add-select-trigger custom-dropdown-trigger" data-action="toggle-quick-add-dropdown">
        <span class="trigger-text custom-dropdown-selected-text">
          ${triggerText}
        </span>
        <span class="trigger-arrow custom-dropdown-arrow">▾</span>
      </button>
      ${isDropdownOpen ? `
        <div class="custom-dropdown-panel custom-dropdown-menu" data-stop>
          <div class="custom-dropdown-search-wrap">
            ${ICONS.search}
            <input type="text" class="custom-dropdown-search-input" id="quick-add-dropdown-search" placeholder="Type party name, bank or acct..." value="${escapeHtml(dropdownSearchVal)}" autocomplete="off">
          </div>
          <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 14px; background:#f8fafc; border-bottom:1px solid var(--border); font-size:12px;">
            <span style="color:var(--secondary-text);">${selectedParties.length} of ${availableParties.length} selected</span>
            <div style="display:flex; gap:12px;">
              <button type="button" class="btn-link" data-action="quick-add-select-all-filtered" style="font-size:11.5px; color:var(--accent); cursor:pointer; background:none; border:none; padding:0; font-weight:600;">
                ${allFilteredSelected ? 'Deselect All' : 'Select All Visible'}
              </button>
              ${selectedParties.length > 0 ? `
                <button type="button" class="btn-link" data-action="quick-add-clear-selection" style="font-size:11.5px; color:var(--danger); cursor:pointer; background:none; border:none; padding:0; font-weight:600;">
                  Clear
                </button>
              ` : ''}
            </div>
          </div>
          <div class="custom-dropdown-options-list">
            ${customDropdownOptionsHtml}
          </div>
        </div>
      ` : ''}
    </div>
  `;

  return `
    <div class="quick-add-payment-box">
      <div class="quick-add-header">
        <h4>${ICONS.plus} Add Another Payment to this Voucher</h4>
        <button type="button" class="btn btn-sm btn-ghost" data-action="toggle-quick-add-mode">
          ${isDirectory ? 'Switch to Manual Input' : 'Switch to Directory Selection'}
        </button>
      </div>

      ${isDirectory ? `
        <div class="quick-add-directory-mode">
          ${selectedParties.length <= 1 ? `
            <div class="quick-add-row-layout">
              <div class="quick-add-col-payee">
                <label class="quick-add-label">Select Party / Parties from Directory</label>
                ${dropdownHtml}
              </div>
              <div class="quick-add-col-amount">
                <label class="quick-add-label">Amount (₹)</label>
                <div class="currency-input">
                  <span class="currency-symbol">₹</span>
                  <input type="number" id="quick-add-directory-amount" min="1" class="amount-input" placeholder="e.g. 50000" inputmode="numeric" value="${selectedParties.length === 1 ? (qa.amounts[selectedParties[0].id] || '') : ''}">
                </div>
              </div>
              <div class="quick-add-col-btn">
                <button type="button" class="btn btn-primary quick-add-submit" data-action="confirm-quick-add-directory" data-id="${h.id}" ${selectedParties.length === 0 ? 'disabled style="opacity:0.6;"' : ''}>
                  ${ICONS.plus} Add Payment
                </button>
              </div>
            </div>
            ${selectedParties.length === 1 ? `
              <div class="quick-add-details">
                <div class="quick-add-detail-item">
                  <span class="detail-label">Bank Name</span>
                  <span class="detail-val">${escapeHtml(selectedParties[0].bankName.toUpperCase())}</span>
                </div>
                <div class="quick-add-detail-item">
                  <span class="detail-label">Account Number</span>
                  <span class="detail-val mono">${renderAccountWithEye(selectedParties[0].accountNo, selectedParties[0].id)}</span>
                </div>
                <div class="quick-add-detail-item">
                  <span class="detail-label">Bank Location</span>
                  <span class="detail-val">${escapeHtml(selectedParties[0].location.toUpperCase())}</span>
                </div>
                <div class="quick-add-detail-item">
                  <span class="detail-label">IFSC Code</span>
                  <span class="detail-val mono">${escapeHtml(selectedParties[0].ifsc.toUpperCase())}</span>
                </div>
              </div>
            ` : ''}
          ` : `
            <!-- Multi-Select Mode -->
            <div class="quick-add-row-layout" style="align-items:flex-end;">
              <div class="quick-add-col-payee" style="flex:2;">
                <label class="quick-add-label">Select Parties from Directory (${selectedParties.length} selected)</label>
                ${dropdownHtml}
              </div>
              <div class="quick-add-col-amount" style="flex:1;">
                <label class="quick-add-label">Set Amount for All (₹)</label>
                <div class="currency-input">
                  <span class="currency-symbol">₹</span>
                  <input type="number" id="quick-add-bulk-amount" min="1" class="amount-input" placeholder="e.g. 50000" inputmode="numeric" value="${qa.defaultAmount || ''}">
                </div>
              </div>
              <div class="quick-add-col-btn">
                <button type="button" class="btn btn-ghost" data-action="apply-quick-add-bulk-amount" style="white-space:nowrap; height:44px;">
                  Apply to All
                </button>
              </div>
            </div>

            <div style="margin-top:14px; border:1px solid var(--border); border-radius:12px; overflow:hidden; background:#ffffff;">
              <div style="max-height:220px; overflow-y:auto;">
                <table style="width:100%; border-collapse:collapse; font-size:13px;">
                  <thead>
                    <tr style="background:rgba(0,0,0,0.02); border-bottom:1px solid var(--border); text-align:left;">
                      <th style="padding:10px 14px; font-weight:600; color:var(--secondary-text); font-size:11.5px; text-transform:uppercase;">Payee / Party</th>
                      <th style="padding:10px 14px; font-weight:600; color:var(--secondary-text); font-size:11.5px; text-transform:uppercase;">Bank & A/C</th>
                      <th style="padding:10px 14px; font-weight:600; color:var(--secondary-text); font-size:11.5px; text-transform:uppercase; text-align:right; width:160px;">Amount (₹)</th>
                      <th style="width:40px;"></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${selectedParties.map(p => `
                      <tr style="border-bottom:1px solid rgba(0,0,0,0.03);">
                        <td style="padding:10px 14px; font-weight:600; color:var(--primary-text);">${escapeHtml(p.name)}</td>
                        <td style="padding:10px 14px; color:var(--secondary-text); font-size:12px;">${escapeHtml(p.bankName)} · <span class="mono">${renderAccountWithEye(p.accountNo, p.id)}</span></td>
                        <td style="padding:10px 14px; text-align:right;">
                          <div class="currency-input" style="height:36px;">
                            <span class="currency-symbol" style="font-size:12px;">₹</span>
                            <input type="number" class="quick-add-multi-amt-input amount-input" data-party-id="${p.id}" min="1" placeholder="0" style="height:36px; font-size:13px; font-weight:600; text-align:right; padding:0 8px 0 24px;" value="${qa.amounts[p.id] || qa.defaultAmount || ''}">
                          </div>
                        </td>
                        <td style="padding:10px 8px; text-align:center;">
                          <button type="button" class="btn-icon" data-action="remove-quick-add-selected-party" data-id="${p.id}" style="color:var(--danger); border:none; background:none; cursor:pointer; font-size:14px;" title="Remove">✕</button>
                        </td>
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
              </div>
              <div style="display:flex; justify-content:space-between; align-items:center; padding:12px 18px; background:rgba(0,0,0,0.02); border-top:1px solid var(--border);">
                <div style="font-size:13px; color:var(--secondary-text);">
                  Ready to add <b>${selectedParties.length}</b> payee${selectedParties.length === 1 ? '' : 's'} to this voucher
                </div>
                <button type="button" class="btn btn-primary" data-action="confirm-quick-add-multiple" data-id="${h.id}">
                  ${ICONS.plus} Add ${selectedParties.length} Payments to Voucher
                </button>
              </div>
            </div>
          `}
        </div>
      ` : `
        <div class="quick-add-manual-mode">
          <div class="field-grid quick-add-manual-grid">
            <div class="field full">
              <label>Particulars (Payee Name)</label>
              <input type="text" id="quick-add-manual-name" placeholder="e.g. Shree Mahavir Timbers">
            </div>
            <div class="field">
              <label>Bank Name</label>
              <input type="text" id="quick-add-manual-bank" placeholder="e.g. Bank of Maharashtra">
            </div>
            <div class="field">
              <div class="field-header-row">
                <label>Account Number</label>
                <span class="val-badge" id="quick-add-manual-acct-val"></span>
              </div>
              <input type="text" id="quick-add-manual-acct" placeholder="e.g. 2053402589" autocomplete="off">
            </div>
            <div class="field">
              <label>Bank Location</label>
              <input type="text" id="quick-add-manual-loc" placeholder="e.g. Anand">
            </div>
            <div class="field">
              <div class="field-header-row">
                <label>IFSC Code</label>
                <span class="val-badge" id="quick-add-manual-ifsc-val"></span>
              </div>
              <input type="text" id="quick-add-manual-ifsc" placeholder="e.g. MAHB0006456" autocomplete="off" style="text-transform: uppercase;">
            </div>
            <div class="field full" style="display:flex; align-items:center; gap:20px; flex-wrap:wrap; margin-top:6px; padding:12px 16px; background:rgba(0,0,0,0.02); border:1px solid var(--border); border-radius:10px;">
              <label style="display:inline-flex; align-items:center; gap:8px; font-size:13px; font-weight:600; color:var(--primary-text); cursor:pointer; margin:0;">
                <input type="checkbox" id="quick-add-manual-save-dir" checked style="width:16px; height:16px; accent-color:var(--accent); cursor:pointer;">
                Save this party to Directory / Parties list
              </label>
              <div style="display:inline-flex; align-items:center; gap:8px;">
                <span style="font-size:12px; font-weight:500; color:var(--secondary-text);">Category:</span>
                <select id="quick-add-manual-category" style="height:34px; padding:0 10px; font-size:12.5px; border:1px solid var(--border); border-radius:8px; background:#ffffff; color:var(--primary-text); outline:none; font-family:inherit;">
                  ${PARTY_CATEGORIES.map(c => `<option value="${c.id}">${c.icon} ${c.shortLabel}</option>`).join('')}
                </select>
              </div>
            </div>
            <div class="field full quick-add-manual-amount-row">
              <label>Amount (₹)</label>
              <div class="quick-add-manual-amount-wrap">
                <div class="currency-input">
                  <span class="currency-symbol">₹</span>
                  <input type="number" id="quick-add-manual-amount" min="1" class="amount-input" placeholder="0" inputmode="numeric">
                </div>
                <button type="button" class="btn btn-primary" data-action="confirm-quick-add-manual" data-id="${h.id}">
                  ${ICONS.plus} Add Payment
                </button>
              </div>
            </div>
          </div>
        </div>
      `}
    </div>
  `;
}

function renderViewHistoryModal(h) {
  const mailSubject = encodeURIComponent(`Combined Cheque Payment Voucher - Cheque No. ${h.chequeNo}`);
  const mailBody = encodeURIComponent(
    `Dear Sir/Madam,\n\n` +
    `Please find attached the combined cheque payment split details sheet for the following transaction:\n\n` +
    `- Payer Bank A/C: ${(h.account ? h.account.holderName : '').toUpperCase()}\n` +
    `- Bank Name: ${(h.account ? h.account.bankName : '').toUpperCase()}\n` +
    `- Account Number: ${h.account ? h.account.accountNo : ''}\n` +
    `- Cheque No: ${h.chequeNo}\n` +
    `- Date: ${formatDateDDMMYYYY(h.date)}\n` +
    `- Total Amount: Rs. ${formatINR(h.total)} /-\n` +
    `- Amount in Words: ${amountToWordsLine(h.total, h.prefix ?? 'INT ')}\n\n` +
    `[IMPORTANT: Please attach the downloaded file: RTGS-NEFT PAYMENT LIST ${formatDateDDMMYYYY(h.date)}.xlsx to this email before sending]\n\n` +
    `Best regards,\n` +
    `${(h.account ? h.account.holderName : '').toUpperCase()}`
  );
  const mailTo = h.account ? h.account.bankEmail : '';
  
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:960px;" data-stop>
        <div class="modal-head">
          <h3>Payment Voucher Preview — ${escapeHtml(h.date)}</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body" style="overflow-x:auto;">
          
          <!-- Excel Sheet Lookalike Table -->
          <div class="excel-preview-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th style="width:60px; text-align:center;">SR.NO.</th>
                  <th>PARTICULARS</th>
                  <th>BANK'S NAME</th>
                  <th>BANK ACCOUNT NUMBER</th>
                  <th>BANK'S LOCATION</th>
                  <th>IFSC CODE</th>
                  <th style="width:140px; text-align:right;">AMOUNT</th>
                  <th style="width:50px;"></th>
                </tr>
              </thead>
              <tbody>
                ${h.parties.map((p, i) => `
                  <tr>
                    <td style="text-align:center; font-weight:bold;">${i + 1}</td>
                    <td>${escapeHtml(p.name.toUpperCase())}</td>
                    <td>${escapeHtml(p.bankName.toUpperCase())}</td>
                    <td class="mono">${renderAccountWithEye(p.accountNo, 'h_modal_' + h.id + '_' + i)}</td>
                    <td>${escapeHtml(p.location.toUpperCase())}</td>
                    <td class="mono">${escapeHtml(p.ifsc.toUpperCase())}</td>
                    <td style="text-align:right;" class="num">
                      <input type="number" class="excel-amt-input" value="${p.amount}" data-action="quick-edit-amount" data-history-id="${h.id}" data-index="${i}" data-stop style="width: 100%; border: none; background: transparent; text-align: right; font-family: inherit; font-weight: bold; font-size: inherit; outline: none; padding: 0; color: #000000;">
                    </td>
                    <td style="text-align:center;">
                      <button class="icon-btn danger" data-action="quick-remove-payment" data-history-id="${h.id}" data-index="${i}" title="Remove Payment" style="width:28px; height:28px; padding:0;">
                        ${ICONS.trash}
                      </button>
                    </td>
                  </tr>
                `).join('')}
                
                <!-- Blank padding rows like Excel if less than 5 rows -->
                ${Array.from({ length: Math.max(0, 5 - h.parties.length) }).map(() => `
                  <tr class="pad-row">
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                `).join('')}

                <!-- Amount in Words & Total merged row -->
                <tr class="total-row">
                  <td>&nbsp;</td>
                  <td colspan="5" class="words-cell">
                    ${escapeHtml(amountToWordsLine(h.total, h.prefix ?? 'INT '))}
                  </td>
                  <td style="text-align:right;" class="total-cell num">
                    ₹ ${formatINR(h.total)}
                  </td>
                  <td>&nbsp;</td>
                </tr>
              </tbody>
            </table>

            <!-- Cheque No & Date section matching Excel placement -->
            <div class="excel-meta-section">
              <div class="meta-item"><b>CHEQUE NO-</b> ${escapeHtml(h.chequeNo)}</div>
              <div class="meta-item"><b>DATE :</b> ${formatDateDDMMYYYY(h.date)}</div>
            </div>
          </div>

          ${renderQuickAddSection(h)}

          <div class="form-actions">
            <button class="btn btn-ghost" data-action="close-modal">Close</button>
            <a href="mailto:${escapeHtml(mailTo)}?subject=${mailSubject}&body=${mailBody}" class="btn btn-ghost" style="text-decoration:none; display:inline-flex; align-items:center; gap:8px;">
              ${ICONS.mail} Draft Email
            </a>
            <button class="btn btn-accent" data-action="redownload-history" data-id="${h.id}">${ICONS.download} Download Excel</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderRunSuccessModal(record) {
  const isEditing = record.editingHistoryId !== null;
  const payerBank = record.account ? record.account.bankName : '—';
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal modal-success" data-stop>
        <div class="modal-body success-modal-body">
          <button class="modal-close success-close" data-action="close-modal" aria-label="Close">${ICONS.x}</button>

          <div class="success-icon-wrap">
            <div class="success-icon-ring">${ICONS.check}</div>
          </div>

          <h3 class="success-title">${isEditing ? 'Payment Run Updated' : 'Payment Run Saved'}</h3>
          <p class="success-subtitle">Your combined cheque voucher is now stored in History Logs.</p>

          <div class="success-summary">
            <div class="success-summary-row">
              <span>Cheque No.</span>
              <strong class="mono">${escapeHtml(record.chequeNo)}</strong>
            </div>
            <div class="success-summary-row">
              <span>Date</span>
              <strong>${escapeHtml(formatDateDDMMYYYY(record.date))}</strong>
            </div>
            <div class="success-summary-row">
              <span>Drawn On</span>
              <strong>${escapeHtml(payerBank)}</strong>
            </div>
            <div class="success-summary-row">
              <span>Payees</span>
              <strong>${record.parties.length} recipient${record.parties.length === 1 ? '' : 's'}</strong>
            </div>
            <div class="success-summary-row highlight">
              <span>Total Amount</span>
              <strong class="num">₹ ${formatINR(record.total)}</strong>
            </div>
          </div>

          <div class="success-actions">
            <button class="btn btn-accent" data-action="close-modal">Add Another Payment</button>
            <button class="btn btn-primary" data-action="goto-history-tab">View in History</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function formatDateDDMMYYYY(isoDate) {
  if (!isoDate) return '';
  const [y, m, d] = isoDate.split('-');
  return `${d}-${m}-${y}`;
}

function renderImportModal() {
  const currentCat = (state.directoryCategory && state.directoryCategory !== 'all') ? state.directoryCategory : 'material';
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:540px;" data-stop>
        <div class="modal-head">
          <h3>Import Parties from Excel</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <div style="margin-bottom:16px;">
            <label style="font-size:12.5px; font-weight:600; color:var(--secondary-text); display:block; margin-bottom:6px;">Assign to Directory Category / List:</label>
            <select id="import-category-select" style="width:100%; height:42px; border:1.5px solid var(--border); border-radius:10px; padding:0 12px; font-family:inherit; font-size:13.5px; background:#ffffff; color:var(--primary-text); outline:none;">
              ${PARTY_CATEGORIES.map(c => `
                <option value="${c.id}" ${c.id === currentCat ? 'selected' : ''}>${c.icon} ${c.label}</option>
              `).join('')}
            </select>
          </div>
          <div class="uploader-box" id="drop-zone" onclick="document.getElementById('excel-file-input').click()">
            ${ICONS.upload}
            <p>Drag & drop your existing bank sheet here or <b>browse files</b></p>
            <input type="file" id="excel-file-input" accept=".xlsx, .xls">
          </div>
          <div id="import-preview-box" style="display:none; margin-top:16px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
              <h4 style="margin:0;font-weight:600;font-size:14px;">Parsed Payees Preview</h4>
              <span id="import-preview-cat-badge" style="font-size:11.5px; font-weight:600; color:var(--accent);"></span>
            </div>
            <div style="max-height:180px; overflow-y:auto; border:1px solid var(--border); border-radius:8px; margin-bottom:16px;" id="import-preview-list">
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center;">
              <span class="hint" id="parsed-count-lbl"></span>
              <button class="btn btn-primary" id="commit-import-btn">Confirm Import</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderBackupModal() {
  return `
    <div class="modal-backdrop" data-action="close-modal">
      <div class="modal" style="max-width:480px;" data-stop>
        <div class="modal-head">
          <h3>Backup / Restore Data</h3>
          <button class="modal-close" data-action="close-modal">${ICONS.x}</button>
        </div>
        <div class="modal-body">
          <div style="margin-bottom:24px;">
            <h4 style="margin:0 0 8px; font-size:14px; font-weight:600;">Export Settings Backup</h4>
            <p style="font-size:13px; color:var(--slate); margin:0 0 14px;">Save a backup of all payee parties, paying accounts, and logs to your computer as a JSON file.</p>
            <button class="btn btn-primary" data-action="backup-export">Download Backup File</button>
          </div>
          <hr style="border:none; border-top:1px solid var(--border); margin:20px 0;">
          <div>
            <h4 style="margin:0 0 8px; font-size:14px; font-weight:600;">Restore Settings Backup</h4>
            <p style="font-size:13px; color:var(--slate); margin:0 0 14px;">Restore settings from a previously saved backup file. This will merge payees and accounts with your current workspace.</p>
            <button class="btn btn-ghost" onclick="document.getElementById('backup-file-input').click()">Upload Backup File</button>
            <input type="file" id="backup-file-input" accept=".json" style="display:none;">
          </div>
        </div>
      </div>
    </div>
  `;
}

function showToast(msg, type) {
  const root = document.getElementById('toast-root');
  if (!root) return;

  let toastType = type;
  if (!toastType) {
    if (/exist|already|duplicate|warning|skipped/i.test(msg)) {
      toastType = 'warning';
    } else if (/error|failed|invalid|cannot|unable/i.test(msg)) {
      toastType = 'error';
    } else {
      toastType = 'success';
    }
  }

  let iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 6 9 17l-5-5"/></svg>`;
  if (toastType === 'warning') {
    iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`;
  } else if (toastType === 'error') {
    iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>`;
  }

  root.innerHTML = `<div class="toast toast-${toastType}">
    ${iconSvg}
    <span>${escapeHtml(msg)}</span>
  </div>`;
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => { root.innerHTML = ''; }, 3200);
}

function setupFormValidation(app) {
  const updateBadge = (badgeEl, status, text) => {
    if (!badgeEl) return;
    badgeEl.className = 'val-badge show ' + status;
    if (status === 'loading') {
      badgeEl.innerHTML = `<span class="loader-spinner-inline"></span> ${text}`;
    } else {
      badgeEl.innerText = text;
    }
  };

  const clearBadge = (badgeEl) => {
    if (!badgeEl) return;
    badgeEl.className = 'val-badge';
    badgeEl.innerHTML = '';
  };

  const setInputHighlight = (inputEl, status) => {
    if (!inputEl) return;
    inputEl.classList.remove('val-input-highlight', 'is-valid', 'is-invalid');
    if (status) {
      inputEl.classList.add('val-input-highlight', 'is-' + status);
    }
  };

  const setupAccountNoField = (inputSelector, badgeSelector) => {
    const input = app.querySelector(inputSelector);
    const badge = app.querySelector(badgeSelector);
    if (!input) return;

    const runValidation = () => {
      const cleanVal = input.value.replace(/\D/g, '');
      if (cleanVal.length === 0) {
        clearBadge(badge);
        setInputHighlight(input, null);
        return;
      }
      if (cleanVal.length >= 9 && cleanVal.length <= 18) {
        updateBadge(badge, 'valid', '✓ Account No. Length Valid');
        setInputHighlight(input, 'valid');
      } else {
        updateBadge(badge, 'invalid', '✗ Invalid Length (9-18 digits)');
        setInputHighlight(input, 'invalid');
      }
    };

    runValidation();

    input.addEventListener('input', () => {
      const rawVal = input.value;
      const cleanVal = rawVal.replace(/\D/g, '');
      if (rawVal !== cleanVal) {
        input.value = cleanVal;
      }
      runValidation();
    });
  };

  const setupIfscField = (inputSelector, badgeSelector, bankInputSelector, locInputSelector) => {
    const input = app.querySelector(inputSelector);
    const badge = app.querySelector(badgeSelector);
    const bankInput = app.querySelector(bankInputSelector);
    const locInput = app.querySelector(locInputSelector);
    if (!input) return;

    const runValidation = () => {
      const val = input.value.trim().toUpperCase();
      if (val.length === 0) {
        clearBadge(badge);
        setInputHighlight(input, null);
        return false;
      }
      const isPatternValid = /^[A-Z]{4}0[A-Z0-9]{6}$/.test(val);
      if (val.length < 11) {
        updateBadge(badge, 'invalid', '✗ Must be 11 characters');
        setInputHighlight(input, 'invalid');
        return false;
      }
      if (!isPatternValid) {
        updateBadge(badge, 'invalid', '✗ Invalid Format');
        setInputHighlight(input, 'invalid');
        return false;
      }
      updateBadge(badge, 'valid', '✓ Pattern Valid');
      setInputHighlight(input, 'valid');
      return true;
    };

    const hasValue = input.value.trim() !== '';
    if (hasValue) {
      runValidation();
    }

    let lastFetched = '';

    input.addEventListener('input', () => {
      let code = input.value.trim().toUpperCase();
      if (code.length > 11) {
        code = code.substring(0, 11);
        input.value = code;
      }
      const isValidPattern = runValidation();
      if (isValidPattern && code !== lastFetched) {
        lastFetched = code;
        updateBadge(badge, 'loading', 'Checking IFSC...');
        setInputHighlight(input, null);
        fetch(`https://ifsc.razorpay.com/${code}`)
          .then(res => {
            if (!res.ok) throw new Error();
            return res.json();
          })
          .then(data => {
            updateBadge(badge, 'valid', '✓ Verified IFSC');
            setInputHighlight(input, 'valid');
            if (bankInput && (!bankInput.value || bankInput.value.trim() === '')) {
              bankInput.value = data.BANK;
              bankInput.dispatchEvent(new Event('input'));
            }
            if (locInput && (!locInput.value || locInput.value.trim() === '')) {
              locInput.value = `${data.BRANCH}, ${data.CENTRE || ''}`.replace(/,\s*$/, '');
              locInput.dispatchEvent(new Event('input'));
            }
          })
          .catch(() => {
            updateBadge(badge, 'invalid', '✗ IFSC Not Found');
            setInputHighlight(input, 'invalid');
          });
      }
    });
  };

  setupAccountNoField('#party-form [name="accountNo"]', '#party-form-acct-val');
  setupIfscField('#party-form [name="ifsc"]', '#party-form-ifsc-val', '#party-form [name="bankName"]', '#party-form [name="location"]');
  setupAccountNoField('#quick-add-manual-acct', '#quick-add-manual-acct-val');
  setupIfscField('#quick-add-manual-ifsc', '#quick-add-manual-ifsc-val', '#quick-add-manual-bank', '#quick-add-manual-loc');
}

function attachHandlers() {
  const app = document.getElementById('app');
  setupFormValidation(app);

  const openDropdown = app.querySelector('.custom-dropdown-options-list');
  if (openDropdown) {
    const selectedOpt = openDropdown.querySelector('.custom-dropdown-option.selected');
    if (selectedOpt) {
      setTimeout(() => {
        selectedOpt.scrollIntoView({ block: 'nearest' });
      }, 0);
    }
  }

  const qaSearch = app.querySelector('#quick-add-dropdown-search');
  if (qaSearch) {
    qaSearch.addEventListener('input', (e) => {
      const q = e.target.value.trim().toLowerCase();
      state.quickAdd.dropdownSearch = e.target.value;
      const list = app.querySelector('.custom-dropdown-options-list');
      if (list) {
        const options = list.querySelectorAll('.custom-dropdown-option');
        let visibleCount = 0;
        options.forEach(opt => {
          const s = (opt.dataset.search || opt.textContent).toLowerCase();
          const match = !q || s.includes(q);
          opt.style.display = match ? 'flex' : 'none';
          if (match) visibleCount++;
        });
        let emptyEl = list.querySelector('.custom-dropdown-empty');
        if (visibleCount === 0) {
          if (!emptyEl) {
            emptyEl = document.createElement('div');
            emptyEl.className = 'custom-dropdown-empty';
            emptyEl.textContent = 'No matching parties found';
            list.appendChild(emptyEl);
          }
          emptyEl.style.display = 'block';
        } else if (emptyEl) {
          emptyEl.style.display = 'none';
        }
      }
    });
  }

  if (!window.__dropdownOutsideClickListenerRegistered) {
    window.addEventListener('click', (e) => {
      if (state.quickAdd && state.quickAdd.dropdownOpen) {
        if (!e.target.closest('.quick-add-custom-select-container')) {
          state.quickAdd.dropdownOpen = false;
          render();
        }
      }
    });
    window.__dropdownOutsideClickListenerRegistered = true;
  }

  app.querySelectorAll('[data-route]').forEach(btn => {
    btn.addEventListener('click', () => {
      navigateTo(btn.dataset.route);
    });
  });

  app.querySelectorAll('[data-action]').forEach(el => {
    el.addEventListener('click', (e) => {
      const action = el.dataset.action;
      handleAction(action, el, e);
    });
    if (el.getAttribute('role') === 'button') {
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          const action = el.dataset.action;
          handleAction(action, el, e);
        }
      });
    }
  });

  app.querySelectorAll('[data-stop]').forEach(el => {
    el.addEventListener('click', e => e.stopPropagation());
  });

  app.querySelectorAll('[data-bind]').forEach(el => {
    const evt = (el.tagName === 'SELECT' || el.type === 'date') ? 'change' : 'input';
    el.addEventListener(evt, () => {
      const path = el.dataset.bind;
      setByPath(path, el.value);
      
      const focusPath = path;
      render();
      const again = document.querySelector(`[data-bind="${focusPath}"]`);
      if (again && evt === 'input') {
        again.focus();
        again.setSelectionRange(again.value.length, again.value.length);
      }
    });
  });

  app.querySelectorAll('[data-action="set-amount"]').forEach(el => {
    el.addEventListener('input', () => {
      const id = el.dataset.id;
      state.run.amounts[id] = el.value;
      renderPartial();
    });
  });

  app.querySelectorAll('[data-action="quick-edit-amount"]').forEach(el => {
    el.addEventListener('input', () => {
      const historyId = el.dataset.historyId;
      const index = parseInt(el.dataset.index, 10);
      const val = Number(el.value) || 0;
      
      const h = state.history.find(x => x.id === historyId);
      if (h && h.parties[index]) {
        h.parties[index].amount = val;
        h.total = h.parties.reduce((sum, p) => sum + p.amount, 0);
        persistHistory();
        
        // Dynamic DOM updates to avoid losing focus
        const totalCells = document.querySelectorAll('.total-cell');
        totalCells.forEach(cell => {
          cell.innerText = `₹ ${formatINR(h.total)}`;
        });
        const wordsCells = document.querySelectorAll('.words-cell');
        wordsCells.forEach(cell => {
          cell.innerText = amountToWordsLine(h.total, h.prefix ?? 'INT ');
        });
      }
    });

    el.addEventListener('change', () => {
      render();
    });
  });

  // History row click routing
  app.querySelectorAll('.history-row').forEach(row => {
    row.addEventListener('click', (e) => {
      if (e.target.closest('[data-stop]') || e.target.closest('.row-actions') || e.target.closest('a')) {
        return;
      }
      const id = row.dataset.id;
      const h = state.history.find(x => x.id === id);
      showViewHistoryModal(h, true);
    });
  });

  const partyForm = document.getElementById('party-form');
  if (partyForm) {
    partyForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(partyForm);
      const data = {
        name: fd.get('name').trim(),
        bankName: fd.get('bankName').trim(),
        accountNo: fd.get('accountNo').trim(),
        location: fd.get('location').trim(),
        ifsc: fd.get('ifsc').trim().toUpperCase(),
        category: fd.get('category') || 'material',
      };
      const id = fd.get('id');
      const normAcct = data.accountNo.replace(/[\s-]+/g, '').toLowerCase();

      if (id) {
        // Edit existing party - check if another party already has this account number in the same category
        const duplicate = state.parties.some(p => p.id !== id && (p.category || 'material') === data.category && normAcct && (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase() === normAcct);
        if (duplicate) {
          showToast(`Another party with this account number already exists in ${getCategoryMeta(data.category).shortLabel}.`);
          return;
        }
        const idx = state.parties.findIndex(p => p.id === id);
        if (idx > -1) state.parties[idx] = { ...state.parties[idx], ...data };
        showToast('Payee updated successfully');
      } else {
        // Add new party - check if already exists in this category
        const duplicate = state.parties.some(p => (p.category || 'material') === data.category && normAcct && (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase() === normAcct);
        if (duplicate) {
          showToast(`A party with this account number already exists in ${getCategoryMeta(data.category).shortLabel}.`);
          return;
        }
        state.parties.push({ id: uid(), ...data });
        showToast('Payee added successfully');
      }
      persistParties();
      state.modal = null;
      render();
    });
  }

  const accountForm = document.getElementById('account-form');
  if (accountForm) {
    accountForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(accountForm);
      const data = {
        holderName: fd.get('holderName').trim(),
        bankName: fd.get('bankName').trim(),
        accountNo: fd.get('accountNo').trim(),
        bankEmail: fd.get('bankEmail').trim(),
        chequeBookStart: fd.get('chequeBookStart') ? fd.get('chequeBookStart').trim() : '',
        chequeBookEnd: fd.get('chequeBookEnd') ? fd.get('chequeBookEnd').trim() : '',
      };
      const id = fd.get('id');
      if (id) {
        const idx = state.myAccounts.findIndex(a => a.id === id);
        if (idx > -1) state.myAccounts[idx] = { ...state.myAccounts[idx], ...data };
        showToast('Paying account updated successfully');
      } else {
        state.myAccounts.push({ id: uid(), ...data });
        showToast('Paying account added successfully');
      }
      persistAccounts();
      state.modal = null;
      render();
    });
  }

  const supabaseForm = document.getElementById('supabase-config-form');
  if (supabaseForm) {
    supabaseForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(supabaseForm);
      const url = fd.get('supabaseUrl') || '';
      const anonKey = fd.get('supabaseAnonKey') || '';
      saveSupabaseConfig(url, anonKey);
      state.supabase.status = 'checking';
      render();
      showToast('Testing Supabase connection...');
      const res = await testSupabaseConnection();
      if (res.success) {
        state.supabase.status = res.needsSchema ? 'needs_schema' : 'connected';
        state.supabase.message = res.message;
        showToast(res.message, res.needsSchema ? 'warning' : 'success');
        if (!res.needsSchema) {
          try {
            const cloudData = await dbPullAllFromCloud();
            if (cloudData) {
              if (cloudData.parties?.length) state.parties = cloudData.parties;
              if (cloudData.myAccounts?.length) state.myAccounts = cloudData.myAccounts;
              if (cloudData.history?.length) state.history = cloudData.history;
              if (cloudData.wordsPrefix) state.wordsPrefix = cloudData.wordsPrefix;
              saveStore(STORE_KEYS.parties, state.parties);
              saveStore(STORE_KEYS.myAccounts, state.myAccounts);
              saveStore(STORE_KEYS.history, state.history);
              state.supabase.lastSync = cloudData.timestamp;
            }
          } catch (err) {
            console.warn('Auto-pull after config notice:', err);
          }
        }
      } else {
        state.supabase.status = 'error';
        state.supabase.message = res.message;
        showToast(`Connection failed: ${res.message}`, 'error');
      }
      render();
    });
  }

  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('excel-file-input');
  if (dropZone && fileInput) {
    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
      }, false);
    });

    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files.length) handleExcelUpload(files[0]);
    });

    fileInput.addEventListener('change', (e) => {
      if (fileInput.files.length) handleExcelUpload(fileInput.files[0]);
    });
  }

  const backupInput = document.getElementById('backup-file-input');
  if (backupInput) {
    backupInput.addEventListener('change', (e) => {
      if (backupInput.files.length) handleBackupRestore(backupInput.files[0]);
    });
  }

  const setupPinForm = document.getElementById('setup-pin-form');
  if (setupPinForm) {
    setupPinForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(setupPinForm);
      const newPin = (fd.get('newPin') || '').trim();
      const confirmPin = (fd.get('confirmPin') || '').trim();
      if (!newPin || !/^\d{4}$/.test(newPin)) {
        showToast('PIN must be exactly 4 numeric digits (0-9)');
        return;
      }
      if (newPin !== confirmPin) {
        showToast('PIN numbers do not match');
        return;
      }
      state.pin = newPin;
      localStorage.setItem(STORE_KEYS.pin, newPin);
      sessionStorage.setItem('pv_session_unlocked', 'true');
      if (isSupabaseConfigured()) {
        dbSaveSetting('master_app_pin', newPin).catch(err => console.warn('PIN cloud sync notice:', err));
      }
      state.modal = null;
      render();
      showToast('4-Digit Security PIN set & synced to cloud');
    });
  }

  const changePinForm = document.getElementById('change-pin-form');
  if (changePinForm) {
    changePinForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(changePinForm);
      const currentPin = (fd.get('currentPin') || '').trim();
      const newPin = (fd.get('newPin') || '').trim();
      const confirmPin = (fd.get('confirmPin') || '').trim();
      if (currentPin !== state.pin) {
        showToast('Incorrect current PIN');
        return;
      }
      if (!newPin || !/^\d{4}$/.test(newPin)) {
        showToast('New PIN must be exactly 4 numeric digits (0-9)');
        return;
      }
      if (newPin !== confirmPin) {
        showToast('New PIN numbers do not match');
        return;
      }
      state.pin = newPin;
      localStorage.setItem(STORE_KEYS.pin, newPin);
      if (isSupabaseConfigured()) {
        dbSaveSetting('master_app_pin', newPin).catch(err => console.warn('PIN cloud sync notice:', err));
      }
      state.modal = null;
      render();
      showToast('PIN changed & synced to cloud');
    });
  }

  const disablePinForm = document.getElementById('disable-pin-form');
  if (disablePinForm) {
    disablePinForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(disablePinForm);
      const currentPin = (fd.get('currentPin') || '').trim();
      if (currentPin !== state.pin) {
        showToast('Incorrect current PIN');
        return;
      }
      state.pin = null;
      localStorage.removeItem(STORE_KEYS.pin);
      sessionStorage.removeItem('pv_session_unlocked');
      if (isSupabaseConfigured()) {
        dbSaveSetting('master_app_pin', null).catch(err => console.warn('PIN cloud sync notice:', err));
      }
      state.modal = null;
      render();
      showToast('PIN protection removed');
    });
  }
}

let tempParsedParties = [];

async function handleExcelUpload(file) {
  try {
    tempParsedParties = await parseExcelParties(file);
    if (!tempParsedParties.length) {
      showToast('No valid payee rows parsed from sheet.');
      return;
    }

    // Smart category suggestion based on filename if not already set by user
    const catSelect = document.getElementById('import-category-select');
    const fname = (file.name || '').toLowerCase();
    if (catSelect) {
      if (/employ|staff|salary|salaries|worker|labour/i.test(fname)) {
        catSelect.value = 'employees';
      } else if (/family|relat|personal|friend|home|house|contract|service|agency/i.test(fname)) {
        catSelect.value = 'family';
      } else if (/utility|bill|rent|office/i.test(fname)) {
        catSelect.value = 'utilities';
      }
    }

    const previewList = document.getElementById('import-preview-list');
    const previewBox = document.getElementById('import-preview-box');
    const countLbl = document.getElementById('parsed-count-lbl');
    const commitBtn = document.getElementById('commit-import-btn');
    const catBadge = document.getElementById('import-preview-cat-badge');

    const refreshImportPreview = () => {
      const selectedCat = catSelect ? catSelect.value : 'material';
      const meta = getCategoryMeta(selectedCat);

      if (catBadge) {
        catBadge.innerHTML = `Assigning to: <b>${meta.icon} ${escapeHtml(meta.shortLabel)}</b>`;
      }

      // Existing accounts in THIS category
      const existingCategoryParties = state.parties.filter(p => (p.category || 'material') === selectedCat);
      const existingAccountSet = new Set(
        existingCategoryParties
          .map(p => (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase())
          .filter(Boolean)
      );
      const existingNameSet = new Set(
        existingCategoryParties.map(p => (p.name || '').trim().toLowerCase()).filter(Boolean)
      );

      const newParties = [];
      const duplicateCategoryParties = [];

      tempParsedParties.forEach(p => {
        const normAcct = (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
        const normName = (p.name || '').trim().toLowerCase();
        const isExistingInCat = normAcct 
          ? existingAccountSet.has(normAcct) 
          : (normName && existingNameSet.has(normName));

        if (isExistingInCat) {
          duplicateCategoryParties.push(p);
        } else {
          newParties.push(p);
        }
      });

      const inSheetDups = tempParsedParties.inSheetDuplicatesCount || 0;

      if (!previewList || !previewBox || !countLbl || !commitBtn) return;

      if (newParties.length === 0) {
        previewList.innerHTML = `
          <div style="padding: 22px 14px; text-align: center; color: var(--secondary-text); font-size: 13.5px; line-height:1.6;">
            ⚠️ <b>All ${tempParsedParties.length} payees</b> in this sheet already exist in <b>${meta.label}</b>.<br>
            <span style="font-size:12px; color: var(--accent);">You can change the target category above or upload a new file.</span>
          </div>
        `;
        countLbl.innerHTML = `<span style="color:var(--accent);">0 new payees (${duplicateCategoryParties.length} existing in ${meta.shortLabel} skipped${inSheetDups > 0 ? `, ${inSheetDups} in-sheet dups ignored` : ''})</span>`;
        commitBtn.disabled = true;
        commitBtn.innerText = 'No New Parties to Import';
      } else {
        commitBtn.disabled = false;
        commitBtn.innerText = `Confirm Import (${newParties.length}) to ${meta.shortLabel}`;

        let html = newParties.map(p => `
          <div style="padding: 10px 14px; border-bottom: 1px solid var(--border); font-size:13px; display:flex; justify-content:space-between; align-items:center;">
            <div>
              <div style="font-weight:600; color:var(--primary-text);">${escapeHtml(p.name)}</div>
              <div style="color:var(--secondary-text); margin-top:2px;">${escapeHtml(p.bankName)} · A/C: ${escapeHtml(p.accountNo)} · IFSC: ${escapeHtml(p.ifsc)}</div>
            </div>
            <span style="font-size:10.5px; padding:2px 8px; border-radius:10px; background:rgba(34,197,94,0.15); color:#16a34a; font-weight:600;">NEW</span>
          </div>
        `).join('');

        if (duplicateCategoryParties.length > 0 || inSheetDups > 0) {
          html += `
            <div style="padding: 10px 14px; background: rgba(239, 68, 68, 0.05); color: var(--secondary-text); font-size: 12px; border-top: 1px dashed var(--border);">
              ${duplicateCategoryParties.length > 0 ? `ℹ️ <b>${duplicateCategoryParties.length} payee${duplicateCategoryParties.length === 1 ? '' : 's'}</b> already exist in ${meta.shortLabel} and will be skipped.<br>` : ''}
              ${inSheetDups > 0 ? `ℹ️ <b>${inSheetDups} duplicate row${inSheetDups === 1 ? '' : 's'}</b> within the Excel sheet were automatically ignored.` : ''}
            </div>
          `;
        }

        previewList.innerHTML = html;
        countLbl.innerHTML = `Found <b>${newParties.length}</b> new payee${newParties.length === 1 ? '' : 's'} ready to import${duplicateCategoryParties.length > 0 ? ` <span style="color:var(--secondary-text);">(${duplicateCategoryParties.length} in ${meta.shortLabel} skipped)</span>` : ''}`;

        commitBtn.onclick = () => {
          let added = 0;
          const currentCatAccounts = new Set(
            state.parties
              .filter(p => (p.category || 'material') === selectedCat)
              .map(p => (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase())
              .filter(Boolean)
          );

          newParties.forEach(tp => {
            const normAcct = (tp.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
            if (!normAcct || !currentCatAccounts.has(normAcct)) {
              if (normAcct) currentCatAccounts.add(normAcct);
              state.parties.push({ id: uid(), category: selectedCat, ...tp });
              added++;
            }
          });
          
          persistParties();
          state.modal = null;
          if (added > 0) {
            showToast(`Imported ${added} new unique payees into ${meta.label}!`);
          } else {
            showToast('No new unique payees were added.');
          }
          render();
        };
      }
      previewBox.style.display = 'block';
    };

    if (catSelect) {
      catSelect.onchange = refreshImportPreview;
    }

    refreshImportPreview();
  } catch (err) {
    showToast(`Error parsing file: ${err.message}`);
  }
}

function handleBackupRestore(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (data && (data.parties || data.myAccounts)) {
        let mergedParties = 0;
        let skippedParties = 0;
        let mergedAccounts = 0;
        
        if (Array.isArray(data.parties)) {
          const currentAccounts = new Set(
            state.parties.map(p => {
              const cat = p.category || 'material';
              const norm = (p.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
              return norm ? `${cat}_${norm}` : `${cat}_name_${(p.name || '').trim().toLowerCase()}`;
            })
          );
          data.parties.forEach(tp => {
            const cat = tp.category || 'material';
            const normAcct = (tp.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
            const key = normAcct ? `${cat}_${normAcct}` : `${cat}_name_${(tp.name || '').trim().toLowerCase()}`;
            if (!currentAccounts.has(key)) {
              currentAccounts.add(key);
              state.parties.push({ id: uid(), category: cat, ...tp });
              mergedParties++;
            } else {
              skippedParties++;
            }
          });
        }
        
        if (Array.isArray(data.myAccounts)) {
          const currentAccounts = new Set(
            state.myAccounts.map(a => (a.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase())
          );
          data.myAccounts.forEach(ta => {
            const normAcct = (ta.accountNo || '').trim().replace(/[\s-]+/g, '').toLowerCase();
            if (normAcct && !currentAccounts.has(normAcct)) {
              currentAccounts.add(normAcct);
              state.myAccounts.push({ id: uid(), ...ta });
              mergedAccounts++;
            }
          });
        }
        
        if (Array.isArray(data.history)) {
          state.history = [...state.history, ...data.history];
          persistHistory();
        }
        
        persistParties();
        persistAccounts();
        state.modal = null;
        showToast(`Restored: ${mergedParties} new parties (${skippedParties} repeated skipped), ${mergedAccounts} accounts.`);
        render();
      } else {
        showToast('Invalid backup file content.');
      }
    } catch (err) {
      showToast('Error reading backup JSON file.');
    }
  };
  reader.readAsText(file);
}

function setByPath(path, value) {
  if (path === 'search') {
    state.search = value;
    return;
  }
  if (path === 'historyPeriod') {
    state.historyPeriod = value;
    return;
  }
  if (path.startsWith('run.')) {
    const key = path.split('.')[1];
    state.run[key] = value;
    if (key === 'prefix') {
      state.wordsPrefix = value;
      persistPrefix();
    }
    return;
  }
  if (path.startsWith('quickAdd.')) {
    const key = path.split('.')[1];
    state.quickAdd[key] = value;
    return;
  }
  if (path.startsWith('monthlyExport.')) {
    const key = path.split('.')[1];
    state.monthlyExport[key] = parseInt(value, 10) || value;
    return;
  }
}

function renderPartial() {
  const selected = state.run.selectedIds
    .map(id => state.parties.find(p => p.id === id))
    .filter(Boolean);
  const total = selected.reduce((sum, p) => sum + (Number(state.run.amounts[p.id]) || 0), 0);

  const totalHead = document.getElementById('amount-sheet-total');
  if (totalHead) {
    totalHead.innerHTML = `Total: <strong>₹ ${formatINR(total)}</strong>`;
  }
  const totalFoot = document.getElementById('amount-sheet-foot-total');
  if (totalFoot) {
    totalFoot.textContent = `₹ ${formatINR(total)}`;
  }

  const strip = document.querySelector('.cheque-strip .words-line');
  if (strip) {
    strip.innerHTML = total > 0
      ? `<span class="words-text">${escapeHtml(amountToWordsLine(total, state.run.prefix))}</span><span class="total-figure">₹ ${formatINR(total)}</span>`
      : `<span class="placeholder">Enter amounts above to generate preview</span>`;
  }
  const genBtn = document.querySelector('[data-action="save-payment-run"]');
  if (genBtn) {
    genBtn.disabled = !(state.run.chequeNo && state.run.date && state.run.accountId && total > 0);
  }
}

function updateLockScreenDOM() {
  const dotsContainer = document.querySelector('.pin-dots-container');
  const msgArea = document.querySelector('.pin-msg-area');
  if (!dotsContainer) return;

  const dots = dotsContainer.querySelectorAll('.pin-dot');
  dots.forEach((dot, idx) => {
    if (idx < state.pinInput.length) {
      dot.classList.add('filled');
    } else {
      dot.classList.remove('filled');
    }
  });

  if (state.pinSuccess) {
    dotsContainer.classList.remove('shake', 'error');
    dotsContainer.classList.add('success');
  } else if (state.pinError) {
    dotsContainer.classList.remove('success');
    dotsContainer.classList.add('shake', 'error');
  } else {
    dotsContainer.classList.remove('shake', 'error', 'success');
  }

  if (msgArea) {
    if (state.pinError) {
      msgArea.innerHTML = `<div class="pin-error-msg">${escapeHtml(state.pinError)}</div>`;
    } else {
      msgArea.innerHTML = `<div class="pin-hint-msg">Tap keypad or type digits on keyboard</div>`;
    }
  }
}

function handlePinDigit(num) {
  if (!state.isLocked || state.pinInput.length >= 4) return;
  state.pinInput += String(num);
  state.pinError = '';
  updateLockScreenDOM();

  if (state.pinInput.length === 4) {
    if (state.pinInput === state.pin) {
      state.pinSuccess = true;
      updateLockScreenDOM();
      
      const lockCard = document.querySelector('.lock-card');
      if (lockCard) {
        lockCard.style.transition = 'all 0.28s cubic-bezier(0.4, 0, 0.2, 1)';
        lockCard.style.opacity = '0';
        lockCard.style.transform = 'scale(0.95) translateY(-8px)';
      }

      setTimeout(() => {
        sessionStorage.setItem('pv_session_unlocked', 'true');
        state.isLocked = false;
        state.pinInput = '';
        state.pinSuccess = false;
        state.pinError = '';
        render();
        showToast('App unlocked');
      }, 250);
    } else {
      state.pinError = 'Incorrect Passcode. Try again.';
      updateLockScreenDOM();
      setTimeout(() => {
        state.pinInput = '';
        state.pinError = '';
        updateLockScreenDOM();
      }, 550);
    }
  }
}

function handlePinBackspace() {
  if (!state.isLocked || state.pinInput.length === 0) return;
  state.pinInput = state.pinInput.slice(0, -1);
  state.pinError = '';
  updateLockScreenDOM();
}

function handlePinClear() {
  if (!state.isLocked) return;
  state.pinInput = '';
  state.pinError = '';
  updateLockScreenDOM();
}

function handleAction(action, el, e) {
  const id = el.dataset.id;

  switch (action) {
    case 'toggle-global-privacy':
      state.privacyMode = !state.privacyMode;
      localStorage.setItem(STORE_KEYS.privacyMode, state.privacyMode ? 'true' : 'false');
      if (state.privacyMode) state.unmaskedIds.clear();
      render();
      showToast(state.privacyMode ? 'Privacy Mode: Account numbers masked' : 'Privacy Mode: Full numbers visible');
      break;

    case 'toggle-unmask-item':
      if (e) e.stopPropagation();
      if (id) {
        if (state.unmaskedIds.has(id)) {
          state.unmaskedIds.delete(id);
        } else {
          state.unmaskedIds.add(id);
        }
        render();
      }
      break;

    case 'open-pin-settings':
      openModal({ type: 'pin-settings' });
      break;

    case 'lock-app':
      if (state.pin) {
        sessionStorage.removeItem('pv_session_unlocked');
        state.isLocked = true;
        state.modal = null;
        state.pinInput = '';
        state.pinError = '';
        render();
        showToast('App locked');
      } else {
        openModal({ type: 'pin-settings' });
      }
      break;

    case 'pin-digit':
      if (el.dataset.num !== undefined) handlePinDigit(el.dataset.num);
      break;

    case 'pin-backspace':
      handlePinBackspace();
      break;

    case 'pin-clear':
      handlePinClear();
      break;

    case 'forgot-pin':
      openModal({ type: 'confirm-reset-pin' });
      break;

    case 'reset-pin-confirmed':
      state.pin = null;
      localStorage.removeItem(STORE_KEYS.pin);
      sessionStorage.removeItem('pv_session_unlocked');
      if (isSupabaseConfigured()) {
        dbSaveSetting('master_app_pin', null).catch(err => console.warn('PIN cloud sync notice:', err));
      }
      state.isLocked = false;
      state.modal = null;
      state.pinInput = '';
      state.pinError = '';
      render();
      showToast('Security PIN has been reset');
      break;

    case 'close-modal':
      closeModal();
      break;

    case 'open-party-form':
      openModal({ type: 'party-form', payload: null }); break;
    case 'view-party-history': {
      const name = el.dataset.name;
      openModal({ type: 'party-history', payload: name }); break;
    }
    case 'edit-party': {
      const p = state.parties.find(x => x.id === id);
      openModal({ type: 'party-form', payload: p }); break;
    }
    case 'delete-party': {
      const p = state.parties.find(x => x.id === id);
      openModal({ type: 'confirm-delete-party', payload: p }); break;
    }
    case 'delete-party-confirmed':
      removePartiesByIds([id]);
      closeModal(true);
      showToast('Payee removed from directory');
      break;

    case 'toggle-directory-party': {
      const idx = state.directorySelectedIds.indexOf(id);
      if (idx > -1) state.directorySelectedIds.splice(idx, 1);
      else state.directorySelectedIds.push(id);
      render();
      break;
    }
    case 'toggle-select-all-directory': {
      const visible = getFilteredParties();
      const allSelected = visible.length > 0 && visible.every(p => state.directorySelectedIds.includes(p.id));
      if (allSelected) {
        const visibleIds = new Set(visible.map(p => p.id));
        state.directorySelectedIds = state.directorySelectedIds.filter(pid => !visibleIds.has(pid));
      } else {
        visible.forEach(p => {
          if (!state.directorySelectedIds.includes(p.id)) state.directorySelectedIds.push(p.id);
        });
      }
      render();
      break;
    }
    case 'delete-selected-parties': {
      if (!state.directorySelectedIds.length) return;
      const ids = [...state.directorySelectedIds];
      const names = ids.map(pid => state.parties.find(p => p.id === pid)?.name).filter(Boolean);
      openModal({ type: 'confirm-delete-parties-bulk', payload: { ids, names } });
      break;
    }
    case 'delete-selected-parties-confirmed': {
      const ids = state.modal?.payload?.ids || [];
      if (!ids.length) return;
      const count = ids.length;
      removePartiesByIds(ids);
      closeModal(true);
      showToast(`${count} payee${count === 1 ? '' : 's'} removed from directory`);
      break;
    }

    case 'open-account-form':
      openModal({ type: 'account-form', payload: null }); break;
    case 'edit-account': {
      const a = state.myAccounts.find(x => x.id === id);
      openModal({ type: 'account-form', payload: a }); break;
    }
    case 'delete-account': {
      const a = state.myAccounts.find(x => x.id === id);
      openModal({ type: 'confirm-delete-account', payload: a }); break;
    }
    case 'delete-account-confirmed':
      state.myAccounts = state.myAccounts.filter(a => a.id !== id);
      persistAccounts();
      if (isSupabaseConfigured()) {
        dbDeleteAccount(id).catch(err => console.warn('Supabase delete account notice:', err));
      }
      closeModal(true);
      showToast('Account removed');
      break;

    case 'toggle-party': {
      const idx = state.run.selectedIds.indexOf(id);
      if (idx > -1) {
        state.run.selectedIds.splice(idx, 1);
        delete state.run.amounts[id];
      } else {
        state.run.selectedIds.push(id);
      }
      render();
      break;
    }
    case 'remove-from-run': {
      const idx = state.run.selectedIds.indexOf(id);
      if (idx > -1) {
        state.run.selectedIds.splice(idx, 1);
        delete state.run.amounts[id];
        render();
      }
      break;
    }
    case 'clear-run':
      state.run = {
        selectedIds: [],
        amounts: {},
        chequeNo: '',
        date: new Date().toISOString().slice(0, 10),
        accountId: '',
        prefix: state.wordsPrefix,
        editingHistoryId: null
      };
      render();
      break;

    case 'cancel-editing-run':
      state.run = {
        selectedIds: [],
        amounts: {},
        chequeNo: '',
        date: new Date().toISOString().slice(0, 10),
        accountId: '',
        prefix: state.wordsPrefix,
        editingHistoryId: null
      };
      navigateTo('history');
      break;

    case 'goto-directory': navigateTo('directory'); break;
    case 'goto-newrun': navigateTo('newrun'); break;
    case 'goto-accounts': navigateTo('accounts'); break;
    case 'goto-history': navigateTo('history'); break;
    case 'goto-dashboard': navigateTo('dashboard'); break;

    case 'set-directory-category':
      state.directoryCategory = el.dataset.cat;
      state.directorySelectedIds = [];
      render();
      break;

    case 'set-run-category':
      state.runCategory = el.dataset.cat;
      render();
      break;

    case 'toggle-select-all-run':
    case 'select-all-category-run': {
      const activeRunCat = state.runCategory || 'all';
      const q = state.search.trim().toLowerCase();
      const filteredRunList = state.parties.filter(p => {
        const partyCat = p.category || 'material';
        if (activeRunCat !== 'all' && partyCat !== activeRunCat) return false;
        return !q || p.name.toLowerCase().includes(q) || p.bankName.toLowerCase().includes(q) || p.accountNo.includes(q) || (p.location && p.location.toLowerCase().includes(q));
      });
      const allVisible = filteredRunList.length > 0 && filteredRunList.every(p => state.run.selectedIds.includes(p.id));
      if (allVisible) {
        const visibleIds = new Set(filteredRunList.map(p => p.id));
        state.run.selectedIds = state.run.selectedIds.filter(pid => !visibleIds.has(pid));
        visibleIds.forEach(vid => {
          delete state.run.amounts[vid];
        });
      } else {
        filteredRunList.forEach(p => {
          if (!state.run.selectedIds.includes(p.id)) {
            state.run.selectedIds.push(p.id);
          }
        });
      }
      render();
      break;
    }

    case 'select-form-category': {
      const card = el;
      const form = card.closest('form');
      if (form) {
        form.querySelectorAll('.category-radio-card').forEach(c => c.classList.remove('selected'));
        card.classList.add('selected');
        const radio = card.querySelector('input[type="radio"]');
        if (radio) radio.checked = true;
      }
      break;
    }

    case 'clear-search':
      state.search = '';
      render();
      break;

    case 'select-payer-account': {
      state.run.accountId = id;
      const acct = state.myAccounts.find(a => a.id === id);
      const chq = getChequeBookInfo(acct);
      if ((!state.run.chequeNo || state.run.chequeNo.trim() === '') && chq && chq.nextUnused) {
        state.run.chequeNo = chq.nextUnused;
      }
      render();
      break;
    }

    case 'auto-fill-next-cheque': {
      const nextNo = el.dataset.next;
      if (nextNo) {
        state.run.chequeNo = nextNo;
        render();
        showToast(`Auto-filled next sequential cheque #${nextNo}`);
      }
      break;
    }

    case 'confirm-duplicate-cheque-save': {
      const record = state.modal?.payload?.record;
      if (record) {
        state.modal = null;
        commitSavePaymentRun(record);
      }
      break;
    }

    case 'save-payment-run':
      savePaymentRun();
      break;

    case 'goto-history-tab':
      closeModal(true);
      navigateTo('history');
      break;

    case 'toggle-quick-add-mode':
      state.quickAdd.mode = state.quickAdd.mode === 'manual' ? 'directory' : 'manual';
      render();
      break;

    case 'toggle-quick-add-dropdown':
    case 'toggle-custom-dropdown': {
      state.quickAdd.dropdownOpen = !state.quickAdd.dropdownOpen;
      if (state.quickAdd.dropdownOpen) {
        state.quickAdd.dropdownSearch = '';
      }
      render();
      break;
    }

    case 'toggle-quick-add-party-select': {
      if (el.dataset.disabled === 'true') break;
      if (!state.quickAdd.selectedIds) state.quickAdd.selectedIds = [];
      const idx = state.quickAdd.selectedIds.indexOf(id);
      if (idx > -1) {
        state.quickAdd.selectedIds.splice(idx, 1);
      } else {
        state.quickAdd.selectedIds.push(id);
      }
      state.quickAdd.partyId = state.quickAdd.selectedIds[0] || '';
      render();
      break;
    }

    case 'select-quick-add-party':
    case 'select-custom-dropdown-party': {
      if (el.dataset.disabled === 'true') break;
      state.quickAdd.selectedIds = [id];
      state.quickAdd.partyId = id;
      state.quickAdd.dropdownOpen = false;
      render();
      break;
    }

    case 'quick-add-select-all-filtered': {
      const h = state.modal?.payload;
      if (!h) break;
      const dsQ = (state.quickAdd.dropdownSearch || '').trim().toLowerCase();
      const availableParties = state.parties.filter(p => {
        const alreadyIn = h.parties.some(hp => (hp.accountNo || '').trim().toLowerCase() === (p.accountNo || '').trim().toLowerCase());
        return !alreadyIn;
      });
      const filteredParties = availableParties.filter(p => {
        if (!dsQ) return true;
        return (
          p.name.toLowerCase().includes(dsQ) ||
          p.bankName.toLowerCase().includes(dsQ) ||
          p.accountNo.includes(dsQ) ||
          (p.location && p.location.toLowerCase().includes(dsQ))
        );
      });
      const allFilteredSelected = filteredParties.length > 0 && filteredParties.every(p => (state.quickAdd.selectedIds || []).includes(p.id));
      if (allFilteredSelected) {
        const filteredIds = new Set(filteredParties.map(p => p.id));
        state.quickAdd.selectedIds = (state.quickAdd.selectedIds || []).filter(pid => !filteredIds.has(pid));
      } else {
        if (!state.quickAdd.selectedIds) state.quickAdd.selectedIds = [];
        filteredParties.forEach(p => {
          if (!state.quickAdd.selectedIds.includes(p.id)) {
            state.quickAdd.selectedIds.push(p.id);
          }
        });
      }
      state.quickAdd.partyId = state.quickAdd.selectedIds[0] || '';
      render();
      break;
    }

    case 'quick-add-clear-selection': {
      state.quickAdd.selectedIds = [];
      state.quickAdd.partyId = '';
      render();
      break;
    }

    case 'remove-quick-add-selected-party': {
      state.quickAdd.selectedIds = (state.quickAdd.selectedIds || []).filter(pid => pid !== id);
      state.quickAdd.partyId = state.quickAdd.selectedIds[0] || '';
      render();
      break;
    }

    case 'apply-quick-add-bulk-amount': {
      const bulkAmtInput = document.getElementById('quick-add-bulk-amount');
      const val = bulkAmtInput ? Number(bulkAmtInput.value) : 0;
      if (val > 0) {
        state.quickAdd.defaultAmount = val;
        document.querySelectorAll('.quick-add-multi-amt-input').forEach(inp => {
          inp.value = val;
          const pid = inp.dataset.partyId;
          if (pid) state.quickAdd.amounts[pid] = val;
        });
        showToast(`Applied ₹ ${formatINR(val)} to all selected payees`);
      } else {
        showToast('Please enter a valid amount to apply');
      }
      break;
    }

    case 'confirm-quick-add-multiple': {
      const h = state.history.find(x => x.id === id);
      if (!h) break;
      const selectedParties = (state.quickAdd.selectedIds || []).map(pid => state.parties.find(p => p.id === pid)).filter(Boolean);
      if (!selectedParties.length) {
        showToast('Please select at least one party.');
        return;
      }

      const amountsMap = {};
      let hasInvalidAmount = false;
      document.querySelectorAll('.quick-add-multi-amt-input').forEach(inp => {
        const pid = inp.dataset.partyId;
        const amt = Number(inp.value);
        if (!amt || amt <= 0) {
          hasInvalidAmount = true;
        }
        amountsMap[pid] = amt;
      });

      if (hasInvalidAmount) {
        showToast('Please enter a valid amount (> 0) for every selected payee.');
        return;
      }

      selectedParties.forEach(party => {
        h.parties.push({
          name: party.name,
          bankName: party.bankName,
          accountNo: party.accountNo,
          location: party.location,
          ifsc: party.ifsc,
          amount: amountsMap[party.id] || 0
        });
      });

      h.total = h.parties.reduce((sum, p) => sum + p.amount, 0);
      persistHistory();
      showToast(`Added ${selectedParties.length} payments to voucher.`);
      showViewHistoryModal(h, true);
      break;
    }

    case 'confirm-quick-add-directory': {
      const amtInput = document.getElementById('quick-add-directory-amount');
      const partyId = (state.quickAdd.selectedIds && state.quickAdd.selectedIds[0]) || state.quickAdd.partyId;
      const amount = amtInput ? Number(amtInput.value) : 0;
      
      if (!partyId) {
        showToast('Please select a party.');
        return;
      }
      if (amount <= 0) {
        showToast('Please enter a valid amount.');
        return;
      }
      
      const party = state.parties.find(x => x.id === partyId);
      const h = state.history.find(x => x.id === id);
      if (party && h) {
        h.parties.push({
          name: party.name,
          bankName: party.bankName,
          accountNo: party.accountNo,
          location: party.location,
          ifsc: party.ifsc,
          amount: amount
        });
        h.total = h.parties.reduce((sum, p) => sum + p.amount, 0);
        persistHistory();
        showToast('Payment added to voucher.');
        showViewHistoryModal(h, true);
      }
      break;
    }
    
    case 'confirm-quick-add-manual': {
      const name = document.getElementById('quick-add-manual-name')?.value.trim();
      const bank = document.getElementById('quick-add-manual-bank')?.value.trim();
      const acct = document.getElementById('quick-add-manual-acct')?.value.trim();
      const loc = document.getElementById('quick-add-manual-loc')?.value.trim();
      const ifsc = document.getElementById('quick-add-manual-ifsc')?.value.trim().toUpperCase();
      const amount = Number(document.getElementById('quick-add-manual-amount')?.value) || 0;
      const saveToDir = document.getElementById('quick-add-manual-save-dir')?.checked;
      const category = document.getElementById('quick-add-manual-category')?.value || 'material';
      
      if (!name || !bank || !acct || !loc || !ifsc) {
        showToast('Please fill all payee details.');
        return;
      }
      if (amount <= 0) {
        showToast('Please enter a valid amount.');
        return;
      }

      if (saveToDir) {
        const normAcct = acct.replace(/[\s-]+/g, '').toLowerCase();
        const exists = state.parties.some(p => (p.accountNo || '').replace(/[\s-]+/g, '').toLowerCase() === normAcct);
        if (!exists) {
          state.parties.push({
            id: uid(),
            name,
            bankName: bank,
            accountNo: acct,
            location: loc,
            ifsc,
            category
          });
          persistParties();
        }
      }
      
      const h = state.history.find(x => x.id === id);
      if (h) {
        h.parties.push({
          name, bankName: bank, accountNo: acct, location: loc, ifsc, amount
        });
        h.total = h.parties.reduce((sum, p) => sum + p.amount, 0);
        persistHistory();
        showToast(saveToDir ? 'Payment added & saved to Directory' : 'Payment added to voucher');
        showViewHistoryModal(h, true);
      }
      break;
    }

    case 'quick-remove-payment': {
      const historyId = el.dataset.historyId;
      const index = parseInt(el.dataset.index, 10);
      const h = state.history.find(x => x.id === historyId);
      if (h && h.parties[index]) {
        h.parties.splice(index, 1);
        h.total = h.parties.reduce((sum, p) => sum + p.amount, 0);
        persistHistory();
        showToast('Payment removed from this run');
        showViewHistoryModal(h, false);
      }
      break;
    }

    case 'view-history': {
      const h = state.history.find(x => x.id === id);
      showViewHistoryModal(h, true);
      break;
    }
    case 'delete-history': {
      const h = state.history.find(x => x.id === id);
      openModal({ type: 'confirm-delete-history', payload: h }); break;
    }
    case 'delete-history-confirmed':
      state.history = state.history.filter(h => h.id !== id);
      persistHistory();
      if (isSupabaseConfigured()) {
        dbDeleteHistory(id).catch(err => console.warn('Supabase delete history notice:', err));
      }
      closeModal(true);
      showToast('Log entry removed');
      break;

    case 'redownload-history': {
      const h = state.history.find(x => x.id === id);
      exportToExcel(h, h.prefix ?? 'INT ');
      break;
    }

    case 'edit-history-run': {
      const record = state.history.find(x => x.id === id);
      if (record) {
        state.run = {
          selectedIds: record.parties.map(p => {
            const found = state.parties.find(x => x.accountNo === p.accountNo);
            return found ? found.id : null;
          }).filter(Boolean),
          amounts: record.parties.reduce((acc, p) => {
            const found = state.parties.find(x => x.accountNo === p.accountNo);
            if (found) acc[found.id] = String(p.amount);
            return acc;
          }, {}),
          chequeNo: record.chequeNo,
          date: record.date,
          accountId: state.myAccounts.find(a => a.accountNo === record.account.accountNo)?.id || '',
          prefix: record.prefix || 'INT ',
          editingHistoryId: record.id
        };
        closeModal(true);
        navigateTo('newrun');
        showToast('Loaded run for editing. You can add/remove payees or change amounts.');
      }
      break;
    }

    case 'open-import-modal':
      openModal({ type: 'import-modal' }); break;

    case 'open-backup-modal':
      openModal({ type: 'backup-modal' }); break;

    case 'backup-export':
      exportBackupJSON();
      break;

    case 'open-export-monthly-modal':
      openModal({ type: 'export-monthly-modal' });
      break;

    case 'confirm-export-monthly': {
      const year = state.monthlyExport.year;
      const month = state.monthlyExport.month;
      
      const filtered = state.history.filter(h => {
        if (!h.date) return false;
        const d = new Date(h.date);
        return d.getFullYear() === year && (d.getMonth() + 1) === month;
      });

      if (filtered.length === 0) {
        const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'long' });
        showToast(`No payment runs found for ${monthName} ${year}`);
        return;
      }

      closeModal(true);
      showToast(`Generating monthly Excel sheet...`);
      exportMonthlyReport(filtered, year, month);
      break;
    }

    case 'test-supabase-connection': {
      if (!isSupabaseConfigured()) {
        showToast('Please enter your Supabase Project URL and Anon API key first.', 'warning');
        return;
      }
      state.supabase.status = 'checking';
      render();
      testSupabaseConnection().then(res => {
        if (res.success) {
          state.supabase.status = res.needsSchema ? 'needs_schema' : 'connected';
          state.supabase.message = res.message;
          showToast(res.message, res.needsSchema ? 'warning' : 'success');
        } else {
          state.supabase.status = 'error';
          state.supabase.message = res.message;
          showToast(`Connection failed: ${res.message}`, 'error');
        }
        render();
      }).catch(err => {
        state.supabase.status = 'error';
        state.supabase.message = err.message;
        showToast(`Connection failed: ${err.message}`, 'error');
        render();
      });
      break;
    }

    case 'clear-supabase-config': {
      saveSupabaseConfig('', '');
      state.supabase.status = 'unconfigured';
      state.supabase.message = '';
      showToast('Disconnected from Supabase. Operating in local offline storage.');
      render();
      break;
    }

    case 'sync-push-to-cloud': {
      if (!isSupabaseConfigured()) {
        showToast('Supabase is not configured yet.', 'error');
        return;
      }
      showToast('Uploading local data to Supabase...');
      dbPushAllToCloud({
        parties: state.parties,
        myAccounts: state.myAccounts,
        history: state.history,
        wordsPrefix: state.wordsPrefix
      }).then(res => {
        state.supabase.status = 'connected';
        state.supabase.lastSync = res.timestamp;
        showToast(`Uploaded to Supabase: ${res.parties} payees, ${res.accounts} accounts, ${res.history} vouchers!`, 'success');
        render();
      }).catch(err => {
        state.supabase.status = 'error';
        state.supabase.message = err.message;
        showToast(`Upload failed: ${err.message}`, 'error');
        render();
      });
      break;
    }

    case 'sync-pull-from-cloud': {
      if (!isSupabaseConfigured()) {
        showToast('Supabase is not configured yet.', 'error');
        return;
      }
      showToast('Downloading data from Supabase...');
      dbPullAllFromCloud().then(cloudData => {
        if (cloudData) {
          state.parties = cloudData.parties || [];
          state.myAccounts = cloudData.myAccounts || [];
          state.history = cloudData.history || [];
          if (cloudData.wordsPrefix) state.wordsPrefix = cloudData.wordsPrefix;
          saveStore(STORE_KEYS.parties, state.parties);
          saveStore(STORE_KEYS.myAccounts, state.myAccounts);
          saveStore(STORE_KEYS.history, state.history);
          state.supabase.status = 'connected';
          state.supabase.lastSync = cloudData.timestamp;
          showToast(`Downloaded: ${state.parties.length} payees, ${state.myAccounts.length} accounts, ${state.history.length} vouchers!`, 'success');
          render();
        }
      }).catch(err => {
        state.supabase.status = 'error';
        state.supabase.message = err.message;
        showToast(`Download failed: ${err.message}`, 'error');
        render();
      });
      break;
    }

    case 'copy-sql-schema': {
      const sql = getSupabaseSchemaSql();
      navigator.clipboard.writeText(sql).then(() => {
        showToast('SQL Setup Script copied to clipboard!', 'success');
      }).catch(() => {
        showToast('Failed to copy SQL script. Please select and copy manually.', 'error');
      });
      break;
    }
  }
}

async function savePaymentRun() {
  const selected = state.run.selectedIds
    .map(id => state.parties.find(p => p.id === id))
    .filter(Boolean);
  const account = state.myAccounts.find(a => a.id === state.run.accountId);
  if (!account || selected.length === 0) return;

  const partiesWithAmounts = selected.map(p => ({
    name: p.name, bankName: p.bankName, accountNo: p.accountNo,
    location: p.location, ifsc: p.ifsc,
    amount: Number(state.run.amounts[p.id]) || 0,
  }));
  const total = partiesWithAmounts.reduce((s, p) => s + p.amount, 0);

  const record = {
    id: state.run.editingHistoryId || uid(),
    createdAt: Date.now(),
    date: state.run.date,
    chequeNo: state.run.chequeNo,
    prefix: state.run.prefix,
    parties: partiesWithAmounts,
    total,
    account: { 
      holderName: account.holderName, 
      bankName: account.bankName, 
      accountNo: account.accountNo,
      bankEmail: account.bankEmail || ''
    },
    editingHistoryId: state.run.editingHistoryId
  };

  // Check for duplicate cheque number in history
  const dupRecord = findDuplicateCheque(state.run.chequeNo, state.run.editingHistoryId);
  if (dupRecord && !state.run.editingHistoryId) {
    openModal({
      type: 'confirm-duplicate-cheque',
      payload: { record, dupRecord }
    });
    return;
  }

  commitSavePaymentRun(record);
}

function commitSavePaymentRun(record) {
  if (state.run.editingHistoryId) {
    const idx = state.history.findIndex(x => x.id === state.run.editingHistoryId);
    if (idx > -1) {
      state.history[idx] = record;
      showToast('Payment run updated successfully');
    }
  } else {
    state.history.push(record);
    showToast('Payment run saved successfully');
  }
  
  persistHistory();
  
  state.run = {
    selectedIds: [],
    amounts: {},
    chequeNo: '',
    date: new Date().toISOString().slice(0, 10),
    accountId: '',
    prefix: state.wordsPrefix,
    editingHistoryId: null
  };
  
  openModal({ type: 'run-success', payload: record });
}

function exportBackupJSON() {
  const backupData = {
    parties: state.parties,
    myAccounts: state.myAccounts,
    history: state.history,
  };
  const json = JSON.stringify(backupData, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Payment-Backup-${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('Backup JSON downloaded successfully');
}

// Prevent mousewheel / trackpad scrolling from incrementing or decrementing number inputs
document.addEventListener('wheel', (e) => {
  if (document.activeElement && document.activeElement.tagName === 'INPUT' && document.activeElement.type === 'number') {
    document.activeElement.blur();
  }
  if (e.target && e.target.tagName === 'INPUT' && e.target.type === 'number') {
    e.preventDefault();
  }
}, { passive: false });

// Keyboard listener for 4-Digit Lock Screen
window.addEventListener('keydown', (e) => {
  if (state.isLocked) {
    if (e.key >= '0' && e.key <= '9') {
      e.preventDefault();
      handlePinDigit(e.key);
    } else if (e.key === 'Backspace') {
      e.preventDefault();
      handlePinBackspace();
    } else if (e.key === 'Escape' || e.key === 'c' || e.key === 'C') {
      e.preventDefault();
      handlePinClear();
    }
  }
});

async function checkSupabaseStatusOnLoad() {
  if (!isSupabaseConfigured()) {
    state.supabase.status = 'unconfigured';
    return;
  }
  state.supabase.status = 'checking';
  try {
    const res = await testSupabaseConnection();
    if (res.success) {
      state.supabase.status = res.needsSchema ? 'needs_schema' : 'connected';
      state.supabase.message = res.message;
      if (!res.needsSchema) {
        try {
          // Fetch cloud app settings (Master PIN, Prefix, etc.)
          try {
            const settings = await dbFetchSettings();
            if (settings && settings.master_app_pin) {
              state.pin = settings.master_app_pin;
              localStorage.setItem(STORE_KEYS.pin, settings.master_app_pin);
              if (sessionStorage.getItem('pv_session_unlocked') !== 'true') {
                state.isLocked = true;
              }
            }
          } catch (e) {
            console.warn('Cloud settings fetch notice:', e);
          }

          const cloudData = await dbPullAllFromCloud();
          const hasCloudData = (cloudData.parties && cloudData.parties.length > 0) ||
                               (cloudData.myAccounts && cloudData.myAccounts.length > 0) ||
                               (cloudData.history && cloudData.history.length > 0);

          if (hasCloudData) {
            // Cloud has data -> hydrate local state
            if (cloudData.parties?.length) state.parties = cloudData.parties;
            if (cloudData.myAccounts?.length) state.myAccounts = cloudData.myAccounts;
            if (cloudData.history?.length) state.history = cloudData.history;
            if (cloudData.wordsPrefix) state.wordsPrefix = cloudData.wordsPrefix;
            saveStore(STORE_KEYS.parties, state.parties);
            saveStore(STORE_KEYS.myAccounts, state.myAccounts);
            saveStore(STORE_KEYS.history, state.history);
            state.supabase.lastSync = cloudData.timestamp;
          } else if (state.parties.length > 0 || state.myAccounts.length > 0 || state.history.length > 0) {
            // Cloud is empty but local storage has data -> automatically upload initial data to Supabase!
            await dbPushAllToCloud({
              parties: state.parties,
              myAccounts: state.myAccounts,
              history: state.history,
              wordsPrefix: state.wordsPrefix
            });
            state.supabase.lastSync = new Date().toISOString();
          }
        } catch (e) {
          console.warn('Initial cloud sync notice:', e);
        }
      }
    } else {
      state.supabase.status = 'error';
      state.supabase.message = res.message;
    }
  } catch (err) {
    state.supabase.status = 'error';
    state.supabase.message = err.message;
  }
  render();
}

render();
checkSupabaseStatusOnLoad();
