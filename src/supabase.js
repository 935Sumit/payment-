import { createClient } from '@supabase/supabase-js';

const CONFIG_KEYS = {
  url: 'pv_supabase_url',
  anonKey: 'pv_supabase_anon_key',
  lastSync: 'pv_last_cloud_sync'
};

let clientInstance = null;

export function getSupabaseConfig() {
  const url = localStorage.getItem(CONFIG_KEYS.url) || (typeof import.meta !== 'undefined' && import.meta.env ? import.meta.env.VITE_SUPABASE_URL : '') || '';
  const anonKey = localStorage.getItem(CONFIG_KEYS.anonKey) || (typeof import.meta !== 'undefined' && import.meta.env ? import.meta.env.VITE_SUPABASE_ANON_KEY : '') || '';
  return { url: url.trim(), anonKey: anonKey.trim() };
}

export function saveSupabaseConfig(url, anonKey) {
  const cleanUrl = (url || '').trim();
  const cleanKey = (anonKey || '').trim();

  if (cleanUrl) {
    localStorage.setItem(CONFIG_KEYS.url, cleanUrl);
  } else {
    localStorage.removeItem(CONFIG_KEYS.url);
  }

  if (cleanKey) {
    localStorage.setItem(CONFIG_KEYS.anonKey, cleanKey);
  } else {
    localStorage.removeItem(CONFIG_KEYS.anonKey);
  }

  clientInstance = null; // Reset client instance to force re-creation
}

export function isSupabaseConfigured() {
  const { url, anonKey } = getSupabaseConfig();
  return Boolean(url && anonKey);
}

export function getSupabaseClient() {
  if (clientInstance) return clientInstance;

  const { url, anonKey } = getSupabaseConfig();
  if (!url || !anonKey) return null;

  try {
    clientInstance = createClient(url, anonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true
      }
    });
    return clientInstance;
  } catch (err) {
    console.error('Failed to initialize Supabase client:', err);
    return null;
  }
}

export async function testSupabaseConnection() {
  const client = getSupabaseClient();
  if (!client) {
    return { success: false, message: 'Supabase Project URL or Anon API Key is missing.' };
  }

  try {
    // Attempt a light query to verify connection
    const { error } = await client.from('app_settings').select('key').limit(1);
    if (error) {
      if (error.code === '42P01') {
        // Relation/table does not exist yet -> connected, but SQL script needs to run
        return { 
          success: true, 
          needsSchema: true, 
          message: 'Connected to Supabase! However, the database tables have not been created yet. Please execute the SQL setup script below in your Supabase SQL Editor.' 
        };
      }
      return { success: false, message: error.message || 'Error communicating with Supabase database.' };
    }
    return { success: true, needsSchema: false, message: 'Connected successfully to Supabase database!' };
  } catch (err) {
    return { success: false, message: err.message || 'Network error connecting to Supabase.' };
  }
}

/**
 * Generates copy-paste SQL schema script for Supabase SQL Editor
 */
export function getSupabaseSchemaSql() {
  return `-- =========================================================
-- PAYMENT VOUCHER GENERATOR - SUPABASE DATABASE SCHEMA
-- Run this script in your Supabase Project > SQL Editor
-- =========================================================

-- 1. App Settings Table (Preferences, Prefix, etc.)
CREATE TABLE IF NOT EXISTS public.app_settings (
    key TEXT PRIMARY KEY,
    value JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 2. Party Directory (Payees)
CREATE TABLE IF NOT EXISTS public.parties (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    bank_name TEXT DEFAULT '',
    account_no TEXT DEFAULT '',
    location TEXT DEFAULT '',
    ifsc TEXT DEFAULT '',
    category TEXT DEFAULT 'material',
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 3. Payer Bank Accounts (My Accounts)
CREATE TABLE IF NOT EXISTS public.my_accounts (
    id TEXT PRIMARY KEY,
    holder_name TEXT NOT NULL,
    bank_name TEXT NOT NULL,
    account_no TEXT NOT NULL,
    bank_email TEXT DEFAULT '',
    cheque_book_start TEXT DEFAULT '',
    cheque_book_end TEXT DEFAULT '',
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 4. Payment History (Voucher Runs)
CREATE TABLE IF NOT EXISTS public.payment_history (
    id TEXT PRIMARY KEY,
    date DATE NOT NULL,
    cheque_no TEXT DEFAULT '',
    prefix TEXT DEFAULT 'INT ',
    parties JSONB NOT NULL DEFAULT '[]'::jsonb,
    total NUMERIC NOT NULL DEFAULT 0,
    account JSONB NOT NULL DEFAULT '{}'::jsonb,
    editing_history_id TEXT,
    created_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Create Fast Lookup Indexes
CREATE INDEX IF NOT EXISTS idx_parties_name ON public.parties (name);
CREATE INDEX IF NOT EXISTS idx_parties_account_no ON public.parties (account_no);
CREATE INDEX IF NOT EXISTS idx_my_accounts_account_no ON public.my_accounts (account_no);
CREATE INDEX IF NOT EXISTS idx_payment_history_date ON public.payment_history (date);
CREATE INDEX IF NOT EXISTS idx_payment_history_cheque ON public.payment_history (cheque_no);

-- Enable Row Level Security (RLS)
ALTER TABLE public.app_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.parties ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.my_accounts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.payment_history ENABLE ROW LEVEL SECURITY;

-- Allow public access with Supabase anon key
DROP POLICY IF EXISTS "Public access to app_settings" ON public.app_settings;
CREATE POLICY "Public access to app_settings" ON public.app_settings FOR ALL USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Public access to parties" ON public.parties;
CREATE POLICY "Public access to parties" ON public.parties FOR ALL USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Public access to my_accounts" ON public.my_accounts;
CREATE POLICY "Public access to my_accounts" ON public.my_accounts FOR ALL USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Public access to payment_history" ON public.payment_history;
CREATE POLICY "Public access to payment_history" ON public.payment_history FOR ALL USING (true) WITH CHECK (true);
`;
}

// -------------------------------------------------------------
// Mapping Helpers (CamelCase <-> SnakeCase)
// -------------------------------------------------------------

function partyToRow(p) {
  return {
    id: p.id,
    name: p.name || '',
    bank_name: p.bankName || '',
    account_no: p.accountNo || '',
    location: p.location || '',
    ifsc: p.ifsc || '',
    category: p.category || 'material',
    updated_at: new Date().toISOString()
  };
}

function rowToParty(r) {
  return {
    id: r.id,
    name: r.name || '',
    bankName: r.bank_name || '',
    accountNo: r.account_no || '',
    location: r.location || '',
    ifsc: r.ifsc || '',
    category: r.category || 'material'
  };
}

function accountToRow(a) {
  return {
    id: a.id,
    holder_name: a.holderName || '',
    bank_name: a.bankName || '',
    account_no: a.accountNo || '',
    bank_email: a.bankEmail || '',
    cheque_book_start: a.chequeBookStart || '',
    cheque_book_end: a.chequeBookEnd || '',
    updated_at: new Date().toISOString()
  };
}

function rowToAccount(r) {
  return {
    id: r.id,
    holderName: r.holder_name || '',
    bankName: r.bank_name || '',
    accountNo: r.account_no || '',
    bankEmail: r.bank_email || '',
    chequeBookStart: r.cheque_book_start || '',
    chequeBookEnd: r.cheque_book_end || ''
  };
}

function historyToRow(h) {
  return {
    id: h.id,
    date: h.date || new Date().toISOString().slice(0, 10),
    cheque_no: h.chequeNo || '',
    prefix: h.prefix || 'INT ',
    parties: Array.isArray(h.parties) ? h.parties : [],
    total: Number(h.total) || 0,
    account: h.account || {},
    editing_history_id: h.editingHistoryId || null,
    created_at: h.createdAt ? new Date(h.createdAt).toISOString() : new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
}

function rowToHistory(r) {
  return {
    id: r.id,
    createdAt: r.created_at ? new Date(r.created_at).getTime() : Date.now(),
    date: r.date,
    chequeNo: r.cheque_no || '',
    prefix: r.prefix || 'INT ',
    parties: Array.isArray(r.parties) ? r.parties : [],
    total: Number(r.total) || 0,
    account: r.account || {},
    editingHistoryId: r.editing_history_id || null
  };
}

// -------------------------------------------------------------
// Database Operations (CRUD)
// -------------------------------------------------------------

export async function dbFetchParties() {
  const client = getSupabaseClient();
  if (!client) return null;
  const { data, error } = await client.from('parties').select('*').order('name', { ascending: true });
  if (error) {
    console.error('Error fetching parties from Supabase:', error);
    throw error;
  }
  return (data || []).map(rowToParty);
}

export async function dbUpsertParty(party) {
  const client = getSupabaseClient();
  if (!client || !party) return;
  const row = partyToRow(party);
  const { error } = await client.from('parties').upsert(row);
  if (error) {
    console.error('Error upserting party to Supabase:', error);
    throw error;
  }
}

export async function dbUpsertParties(partiesList) {
  const client = getSupabaseClient();
  if (!client || !partiesList || partiesList.length === 0) return;
  const rows = partiesList.map(partyToRow);
  const { error } = await client.from('parties').upsert(rows);
  if (error) {
    console.error('Error bulk upserting parties to Supabase:', error);
    throw error;
  }
}

export async function dbDeleteParty(id) {
  const client = getSupabaseClient();
  if (!client || !id) return;
  const { error } = await client.from('parties').delete().eq('id', id);
  if (error) {
    console.error('Error deleting party from Supabase:', error);
    throw error;
  }
}

export async function dbDeleteParties(ids) {
  const client = getSupabaseClient();
  if (!client || !ids || ids.length === 0) return;
  const { error } = await client.from('parties').delete().in('id', ids);
  if (error) {
    console.error('Error bulk deleting parties from Supabase:', error);
    throw error;
  }
}

export async function dbFetchAccounts() {
  const client = getSupabaseClient();
  if (!client) return null;
  const { data, error } = await client.from('my_accounts').select('*').order('bank_name', { ascending: true });
  if (error) {
    console.error('Error fetching accounts from Supabase:', error);
    throw error;
  }
  return (data || []).map(rowToAccount);
}

export async function dbUpsertAccount(account) {
  const client = getSupabaseClient();
  if (!client || !account) return;
  const row = accountToRow(account);
  const { error } = await client.from('my_accounts').upsert(row);
  if (error) {
    console.error('Error upserting account to Supabase:', error);
    throw error;
  }
}

export async function dbUpsertAccounts(accountsList) {
  const client = getSupabaseClient();
  if (!client || !accountsList || accountsList.length === 0) return;
  const rows = accountsList.map(accountToRow);
  const { error } = await client.from('my_accounts').upsert(rows);
  if (error) {
    console.error('Error bulk upserting accounts to Supabase:', error);
    throw error;
  }
}

export async function dbDeleteAccount(id) {
  const client = getSupabaseClient();
  if (!client || !id) return;
  const { error } = await client.from('my_accounts').delete().eq('id', id);
  if (error) {
    console.error('Error deleting account from Supabase:', error);
    throw error;
  }
}

export async function dbFetchHistory() {
  const client = getSupabaseClient();
  if (!client) return null;
  const { data, error } = await client.from('payment_history').select('*').order('date', { ascending: false });
  if (error) {
    console.error('Error fetching history from Supabase:', error);
    throw error;
  }
  return (data || []).map(rowToHistory);
}

export async function dbUpsertHistory(historyItem) {
  const client = getSupabaseClient();
  if (!client || !historyItem) return;
  const row = historyToRow(historyItem);
  const { error } = await client.from('payment_history').upsert(row);
  if (error) {
    console.error('Error upserting history to Supabase:', error);
    throw error;
  }
}

export async function dbUpsertHistories(historyList) {
  const client = getSupabaseClient();
  if (!client || !historyList || historyList.length === 0) return;
  const rows = historyList.map(historyToRow);
  const { error } = await client.from('payment_history').upsert(rows);
  if (error) {
    console.error('Error bulk upserting history to Supabase:', error);
    throw error;
  }
}

export async function dbDeleteHistory(id) {
  const client = getSupabaseClient();
  if (!client || !id) return;
  const { error } = await client.from('payment_history').delete().eq('id', id);
  if (error) {
    console.error('Error deleting history item from Supabase:', error);
    throw error;
  }
}

export async function dbFetchSettings() {
  const client = getSupabaseClient();
  if (!client) return null;
  const { data, error } = await client.from('app_settings').select('*');
  if (error) {
    console.error('Error fetching settings from Supabase:', error);
    return null;
  }
  const settings = {};
  (data || []).forEach(row => {
    settings[row.key] = row.value;
  });
  return settings;
}

export async function dbSaveSetting(key, value) {
  const client = getSupabaseClient();
  if (!client || !key) return;
  const { error } = await client.from('app_settings').upsert({
    key,
    value,
    updated_at: new Date().toISOString()
  });
  if (error) {
    console.error(`Error saving setting '${key}' to Supabase:`, error);
  }
}

// -------------------------------------------------------------
// High-level Push / Pull Sync
// -------------------------------------------------------------

export async function dbPushAllToCloud({ parties, myAccounts, history, wordsPrefix }) {
  const client = getSupabaseClient();
  if (!client) {
    throw new Error('Supabase client is not configured.');
  }

  const results = { parties: 0, accounts: 0, history: 0 };

  if (parties && parties.length > 0) {
    await dbUpsertParties(parties);
    results.parties = parties.length;
  }

  if (myAccounts && myAccounts.length > 0) {
    await dbUpsertAccounts(myAccounts);
    results.accounts = myAccounts.length;
  }

  if (history && history.length > 0) {
    await dbUpsertHistories(history);
    results.history = history.length;
  }

  if (wordsPrefix !== undefined) {
    await dbSaveSetting('wordsPrefix', wordsPrefix);
  }

  const syncTime = new Date().toISOString();
  localStorage.setItem(CONFIG_KEYS.lastSync, syncTime);

  return { ...results, timestamp: syncTime };
}

export async function dbPullAllFromCloud() {
  const client = getSupabaseClient();
  if (!client) {
    throw new Error('Supabase client is not configured.');
  }

  const [parties, myAccounts, history, settings] = await Promise.all([
    dbFetchParties(),
    dbFetchAccounts(),
    dbFetchHistory(),
    dbFetchSettings()
  ]);

  const syncTime = new Date().toISOString();
  localStorage.setItem(CONFIG_KEYS.lastSync, syncTime);

  return {
    parties: parties || [],
    myAccounts: myAccounts || [],
    history: history || [],
    wordsPrefix: settings?.wordsPrefix,
    timestamp: syncTime
  };
}
