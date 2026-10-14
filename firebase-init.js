// firebase-init.js — connects this static site to a shared Firestore
// database so every visitor sees the same, most-recently-uploaded data.
// Loaded as <script type="module"> so we can use Firebase's ES module SDK
// straight from a CDN, with no npm/build step required.

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getFirestore,
  doc,
  setDoc,
  onSnapshot,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyBtY_7oiNaQGgtJR9V_vpbBOxOsO62qnHs",
  authDomain: "cm-maintenance-cost.firebaseapp.com",
  projectId: "cm-maintenance-cost",
  storageBucket: "cm-maintenance-cost.firebasestorage.app",
  messagingSenderId: "1030249075400",
  appId: "1:1030249075400:web:dc3274cf059128bb078bfd",
  measurementId: "G-CYNHSD8Q22",
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Each xlsx source gets its own document so a single update never has to
// rewrite the other source's (larger) payload, and so neither document gets
// anywhere close to Firestore's 1MB-per-document ceiling.
const COLLECTION = "dashboard";

window.__fb = {
  /** Overwrite the shared document for one source ("iw28" | "iw38"). */
  async saveDoc(id, data) {
    await setDoc(doc(db, COLLECTION, id), data);
  },
  /** Subscribe to live updates for one source. Returns an unsubscribe fn.
   *  `cb` fires immediately with the current value (or null) and again
   *  every time any visitor overwrites that document. */
  subscribe(id, cb) {
    return onSnapshot(
      doc(db, COLLECTION, id),
      (snap) => cb(snap.exists() ? snap.data() : null),
      (err) => {
        console.error("[firebase] subscribe error for", id, err);
        cb(null, err);
      }
    );
  },
};

window.dispatchEvent(new Event("fb-ready"));
