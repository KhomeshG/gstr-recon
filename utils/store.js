'use strict';
const crypto = require('crypto');

/** Tiny in-memory result store with expiry (closure-based, no classes). */
const createStore = (ttlMs = 60 * 60 * 1000) => {
  const entries = new Map();
  const sweep = () => {
    const now = Date.now();
    [...entries].filter(([, v]) => v.expires < now).forEach(([k]) => entries.delete(k));
  };
  return {
    save: (data) => {
      sweep();
      const id = crypto.randomBytes(8).toString('hex');
      entries.set(id, { data, expires: Date.now() + ttlMs });
      return id;
    },
    set: (id, data) => {
      const prev = entries.get(id);
      if (prev) entries.set(id, { data, expires: Date.now() + ttlMs });
      return Boolean(prev);
    },
    get: (id) => {
      sweep();
      return entries.get(id)?.data ?? null;
    },
  };
};

module.exports = { createStore };
