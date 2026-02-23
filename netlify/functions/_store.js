const crypto = require("crypto");

const MEMORY_STORE = new Map();
let blobStore = null;
let blobInitAttempted = false;

async function getBlobStore() {
  if (blobInitAttempted) return blobStore;
  blobInitAttempted = true;
  try {
    // Netlify Blobs is available in production; fallback to memory locally.
    const { getStore } = require("@netlify/blobs");
    blobStore = getStore("shay-analytics-ai");
  } catch (error) {
    blobStore = null;
    console.warn("Netlify Blobs not available. Falling back to in-memory store.");
  }
  return blobStore;
}

function memorySet(key, value, ttlSeconds) {
  const expiresAt = Date.now() + ttlSeconds * 1000;
  MEMORY_STORE.set(key, { value, expiresAt });
}

function memoryGet(key) {
  const entry = MEMORY_STORE.get(key);
  if (!entry) return null;
  if (Date.now() > entry.expiresAt) {
    MEMORY_STORE.delete(key);
    return null;
  }
  return entry.value;
}

function memoryDelete(key) {
  MEMORY_STORE.delete(key);
}

async function setItem(key, value, ttlSeconds) {
  const store = await getBlobStore();
  if (store) {
    await store.set(key, value, { ttl: ttlSeconds });
    return;
  }
  memorySet(key, value, ttlSeconds);
}

async function getItem(key) {
  const store = await getBlobStore();
  if (store) {
    const value = await store.get(key, { type: "json" });
    return value || null;
  }
  return memoryGet(key);
}

async function deleteItem(key) {
  const store = await getBlobStore();
  if (store) {
    await store.delete(key);
    return;
  }
  memoryDelete(key);
}

function randomId() {
  return crypto.randomBytes(16).toString("hex");
}

module.exports = {
  setItem,
  getItem,
  deleteItem,
  randomId,
};
