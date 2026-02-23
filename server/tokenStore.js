const path = require("path");
const sqlite3 = require("sqlite3").verbose();

const DB_PATH = path.join(__dirname, "tokens.sqlite");
const memoryStore = new Map();

let db = null;
let dbReady = false;
let dbFailed = false;

function initDb() {
  if (dbReady || dbFailed) return;
  try {
    db = new sqlite3.Database(DB_PATH);
    db.serialize(() => {
      db.run("CREATE TABLE IF NOT EXISTS sessions (id TEXT PRIMARY KEY, tokens TEXT, updatedAt INTEGER)");
      db.run("CREATE TABLE IF NOT EXISTS states (state TEXT PRIMARY KEY, createdAt INTEGER)");
    });
    dbReady = true;
  } catch (error) {
    dbFailed = true;
    console.warn("SQLite unavailable. Falling back to in-memory token store.", error);
  }
}

function useMemory() {
  return dbFailed || !dbReady;
}

async function setSession(id, tokens) {
  initDb();
  if (useMemory()) {
    memoryStore.set(`session:${id}`, { tokens, updatedAt: Date.now() });
    return;
  }
  return new Promise((resolve, reject) => {
    db.run(
      "INSERT OR REPLACE INTO sessions (id, tokens, updatedAt) VALUES (?, ?, ?)",
      [id, JSON.stringify(tokens), Date.now()],
      (err) => (err ? reject(err) : resolve())
    );
  });
}

async function getSession(id) {
  initDb();
  if (useMemory()) {
    return memoryStore.get(`session:${id}`)?.tokens || null;
  }
  return new Promise((resolve, reject) => {
    db.get("SELECT tokens FROM sessions WHERE id = ?", [id], (err, row) => {
      if (err) return reject(err);
      if (!row) return resolve(null);
      try {
        resolve(JSON.parse(row.tokens));
      } catch (error) {
        resolve(null);
      }
    });
  });
}

async function setState(state) {
  initDb();
  if (useMemory()) {
    memoryStore.set(`state:${state}`, { createdAt: Date.now() });
    return;
  }
  return new Promise((resolve, reject) => {
    db.run(
      "INSERT OR REPLACE INTO states (state, createdAt) VALUES (?, ?)",
      [state, Date.now()],
      (err) => (err ? reject(err) : resolve())
    );
  });
}

async function consumeState(state) {
  initDb();
  if (useMemory()) {
    const exists = memoryStore.has(`state:${state}`);
    memoryStore.delete(`state:${state}`);
    return exists;
  }
  return new Promise((resolve, reject) => {
    db.get("SELECT state FROM states WHERE state = ?", [state], (err, row) => {
      if (err) return reject(err);
      db.run("DELETE FROM states WHERE state = ?", [state], () => {
        resolve(!!row);
      });
    });
  });
}

module.exports = {
  setSession,
  getSession,
  setState,
  consumeState,
};
