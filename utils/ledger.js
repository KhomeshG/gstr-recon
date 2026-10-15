'use strict';
/**
 * Tiny JSON-file persistence for the unreconciled-ITC ledger (data/unreconciled_ledger.json).
 * Shape: { pending: [entry], claimed: [entry + claimedIn, claimedAt] }.
 * All ledger *logic* lives in reconcile.js as pure functions; this file only does I/O.
 */
const fs = require('fs');
const path = require('path');
const { emptyLedger } = require('./reconcile');

const createLedgerFile = (file = path.join(__dirname, '..', 'data', 'unreconciled_ledger.json')) => ({
  file,
  load: () => {
    try {
      const data = JSON.parse(fs.readFileSync(file, 'utf8'));
      return { pending: Array.isArray(data.pending) ? data.pending : [], claimed: Array.isArray(data.claimed) ? data.claimed : [] };
    } catch (e) {
      return emptyLedger(); // missing or corrupt file -> start fresh
    }
  },
  save: (ledger) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const tmp = `${file}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(ledger, null, 2));
    fs.renameSync(tmp, file); // atomic swap so a crash never leaves half a file
  },
});

module.exports = { createLedgerFile };
