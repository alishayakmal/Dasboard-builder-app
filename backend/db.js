const path = require("path");
const sqlite3 = require("sqlite3").verbose();

const DB_PATH = path.join(__dirname, "tokens.db");

const db = new sqlite3.Database(DB_PATH);

db.serialize(() => {
  db.run("CREATE TABLE IF NOT EXISTS tokens (sessionId TEXT PRIMARY KEY, refreshToken TEXT, updatedAt INTEGER)");
  db.run("CREATE TABLE IF NOT EXISTS states (state TEXT PRIMARY KEY, createdAt INTEGER)");
});

function setRefreshToken(sessionId, refreshToken) {
  return new Promise((resolve, reject) => {
    db.run(
      "INSERT OR REPLACE INTO tokens (sessionId, refreshToken, updatedAt) VALUES (?, ?, ?)",
      [sessionId, refreshToken, Date.now()],
      (err) => (err ? reject(err) : resolve())
    );
  });
}

function getRefreshToken(sessionId) {
  return new Promise((resolve, reject) => {
    db.get("SELECT refreshToken FROM tokens WHERE sessionId = ?", [sessionId], (err, row) => {
      if (err) return reject(err);
      resolve(row ? row.refreshToken : null);
    });
  });
}

function setState(state) {
  return new Promise((resolve, reject) => {
    db.run("INSERT OR REPLACE INTO states (state, createdAt) VALUES (?, ?)", [state, Date.now()], (err) =>
      err ? reject(err) : resolve()
    );
  });
}

function consumeState(state) {
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
  setRefreshToken,
  getRefreshToken,
  setState,
  consumeState,
};
