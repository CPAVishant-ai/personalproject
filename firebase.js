// =====================================================================
//  firebase.js — Firebase v10 modular SDK (loaded as ES module)
//  Exposes window._fb for use by app.js (non-module global scope)
// =====================================================================

import { initializeApp }   from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-app.js';
import {
  getAuth,
  signInAnonymously,
  signOut,
  onAuthStateChanged,
} from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js';
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  deleteDoc,
  updateDoc,
  collection,
  getDocs,
  writeBatch,
} from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js';

const firebaseConfig = window.FIREBASE_CONFIG;

// ── Init ──────────────────────────────────────────────────────────────
const app  = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db   = getFirestore(app);

// ── Stable owner UID ──────────────────────────────────────────────────
// Anonymous auth generates a new uid each session after signOut().
// We pin the very first uid in localStorage so data is always at the
// same Firestore path regardless of how many times the user signs out/in.
const OWNER_UID_KEY = 'rupaiya_owner_uid';

// ── Helpers ───────────────────────────────────────────────────────────
function userExpenseRef(uid, id)   { return doc(db, 'users', uid, 'expenses', id); }
function expensesCol(uid)          { return collection(db, 'users', uid, 'expenses'); }
function settingsRef(uid)          { return doc(db, 'users', uid, 'data', 'settings'); }

// ── Firestore operations ──────────────────────────────────────────────
async function loadUserData(uid) {
  const [settingsSnap, expSnap] = await Promise.all([
    getDoc(settingsRef(uid)),
    getDocs(expensesCol(uid)),
  ]);
  const settings = settingsSnap.exists() ? settingsSnap.data() : {};
  const expenses = [];
  expSnap.forEach(d => expenses.push({ id: d.id, ...d.data() }));
  return { settings, expenses };
}

async function saveSettings(uid, settings) {
  await setDoc(settingsRef(uid), settings);
}

async function addExpense(uid, expense) {
  const { id, ...data } = expense;
  await setDoc(userExpenseRef(uid, id), data);
}

async function deleteExpense(uid, expenseId) {
  await deleteDoc(userExpenseRef(uid, expenseId));
}

async function updateExpense(uid, expenseId, data) {
  await updateDoc(userExpenseRef(uid, expenseId), data);
}

// Firestore batch limit is 500 — chunk automatically
async function batchAddExpenses(uid, expenses) {
  const CHUNK = 499;
  for (let i = 0; i < expenses.length; i += CHUNK) {
    const batch = writeBatch(db);
    expenses.slice(i, i + CHUNK).forEach(expense => {
      const { id, ...data } = expense;
      batch.set(userExpenseRef(uid, id), data);
    });
    await batch.commit();
  }
}

async function batchDeleteExpenses(uid, ids) {
  const CHUNK = 499;
  for (let i = 0; i < ids.length; i += CHUNK) {
    const batch = writeBatch(db);
    ids.slice(i, i + CHUNK).forEach(id => batch.delete(userExpenseRef(uid, id)));
    await batch.commit();
  }
}

async function batchUpdateExpenses(uid, updates) {
  // updates: Array<{ id: string, data: object }>
  const CHUNK = 499;
  for (let i = 0; i < updates.length; i += CHUNK) {
    const batch = writeBatch(db);
    updates.slice(i, i + CHUNK).forEach(({ id, data }) => {
      batch.update(userExpenseRef(uid, id), data);
    });
    await batch.commit();
  }
}

async function clearAllExpenses(uid) {
  const snap = await getDocs(expensesCol(uid));
  const CHUNK = 499;
  const docs = [];
  snap.forEach(d => docs.push(d.ref));
  for (let i = 0; i < docs.length; i += CHUNK) {
    const batch = writeBatch(db);
    docs.slice(i, i + CHUNK).forEach(ref => batch.delete(ref));
    await batch.commit();
  }
}

// ── Expose to global scope ────────────────────────────────────────────
window._fb = {
  signIn:               () => signInAnonymously(auth),
  signOut:              () => signOut(auth),
  loadUserData,
  saveSettings,
  addExpense,
  deleteExpense,
  updateExpense,
  batchAddExpenses,
  batchDeleteExpenses,
  batchUpdateExpenses,
  clearAllExpenses,
};

// ── Auth state → dispatch custom event for app.js ────────────────────
onAuthStateChanged(auth, user => {
  if (user) {
    // Pin the first anonymous uid so Firestore path never changes
    let ownerUid = localStorage.getItem(OWNER_UID_KEY);
    if (!ownerUid) {
      ownerUid = user.uid;
      localStorage.setItem(OWNER_UID_KEY, ownerUid);
    }
    // Expose a stable synthetic user object (no displayName/photo for anon)
    document.dispatchEvent(new CustomEvent('firebase:authstate', {
      detail: { uid: ownerUid, displayName: 'Owner', email: null, photoURL: null },
    }));
  } else {
    document.dispatchEvent(new CustomEvent('firebase:authstate', { detail: null }));
  }
});
