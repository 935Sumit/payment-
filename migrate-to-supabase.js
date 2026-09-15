/**
 * Standalone Migration Script: Local Data -> Supabase
 * 
 * Instructions:
 * 1. Put your Supabase credentials in the variables below (or pass via environment variables).
 * 2. If you have a backup JSON file exported from the app, put its filename in BACKUP_FILE.
 * 3. Run: node migrate-to-supabase.js
 */

import { createClient } from '@supabase/supabase-js';
import fs from 'fs';
import path from 'path';

// --- CONFIGURATION ---
const SUPABASE_URL = process.env.VITE_SUPABASE_URL || 'YOUR_SUPABASE_PROJECT_URL';
const SUPABASE_ANON_KEY = process.env.VITE_SUPABASE_ANON_KEY || 'YOUR_SUPABASE_ANON_KEY';
const BACKUP_FILE = process.argv[2] || 'backup.json'; // Optional: pass backup file path as argument

async function runMigration() {
  if (SUPABASE_URL === 'YOUR_SUPABASE_PROJECT_URL' || SUPABASE_ANON_KEY === 'YOUR_SUPABASE_ANON_KEY') {
    console.log('\n❌ Please provide your Supabase URL and Anon Key in the script or via environment variables.');
    console.log('Usage:');
    console.log('  node migrate-to-supabase.js <path-to-exported-backup.json>\n');
    console.log('Or use the built-in "🚀 Upload Local to Cloud" button in the Web App interface (Cloud & Supabase tab).\n');
    return;
  }

  const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  let data = null;
  if (fs.existsSync(BACKUP_FILE)) {
    console.log(`\n📂 Reading data from backup file: ${BACKUP_FILE}...`);
    try {
      const raw = fs.readFileSync(BACKUP_FILE, 'utf-8');
      data = JSON.parse(raw);
    } catch (err) {
      console.error('❌ Error reading backup JSON:', err.message);
      return;
    }
  } else {
    console.log(`\n⚠️  Backup file "${BACKUP_FILE}" not found.`);
    console.log('💡 Tip: Export your data from the app via "Party Directory > Backup / Restore > Download Backup File",');
    console.log('   then run: node migrate-to-supabase.js Voucher-Book-Backup-XXXX.json\n');
    console.log('👉 Or simply use the web app: go to "Cloud & Supabase" tab and click "🚀 Upload Local to Cloud".\n');
    return;
  }

  console.log('🚀 Starting migration to Supabase...\n');

  // 1. Parties
  if (Array.isArray(data.parties) && data.parties.length > 0) {
    console.log(`Uploading ${data.parties.length} payees...`);
    const partyRows = data.parties.map(p => ({
      id: p.id,
      name: p.name || '',
      bank_name: p.bankName || '',
      account_no: p.accountNo || '',
      location: p.location || '',
      ifsc: p.ifsc || '',
      category: p.category || 'material',
      updated_at: new Date().toISOString()
    }));

    const { error } = await supabase.from('parties').upsert(partyRows);
    if (error) {
      console.error('❌ Error uploading parties:', error.message);
    } else {
      console.log(`✅ ${data.parties.length} payees successfully uploaded!`);
    }
  }

  // 2. Accounts
  if (Array.isArray(data.myAccounts) && data.myAccounts.length > 0) {
    console.log(`Uploading ${data.myAccounts.length} bank accounts...`);
    const accountRows = data.myAccounts.map(a => ({
      id: a.id,
      holder_name: a.holderName || '',
      bank_name: a.bankName || '',
      account_no: a.accountNo || '',
      bank_email: a.bankEmail || '',
      cheque_book_start: a.chequeBookStart || '',
      cheque_book_end: a.chequeBookEnd || '',
      updated_at: new Date().toISOString()
    }));

    const { error } = await supabase.from('my_accounts').upsert(accountRows);
    if (error) {
      console.error('❌ Error uploading accounts:', error.message);
    } else {
      console.log(`✅ ${data.myAccounts.length} bank accounts successfully uploaded!`);
    }
  }

  // 3. Payment History
  if (Array.isArray(data.history) && data.history.length > 0) {
    console.log(`Uploading ${data.history.length} payment vouchers...`);
    const historyRows = data.history.map(h => ({
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
    }));

    const { error } = await supabase.from('payment_history').upsert(historyRows);
    if (error) {
      console.error('❌ Error uploading payment history:', error.message);
    } else {
      console.log(`✅ ${data.history.length} payment vouchers successfully uploaded!`);
    }
  }

  console.log('\n🎉 Migration process completed!\n');
}

runMigration();
