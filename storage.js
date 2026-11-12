/*
 * مُعبِّئ الاستبيانات — الحفظ داخل المتصفح
 * تُحفظ الاستبيانات (القالب، الإعدادات، آخر ملف إكسل، سجل الطلبات المنجزة)
 * في قاعدة بيانات المتصفح IndexedDB على هذا الجهاز فقط، ولا يُرسل منها شيء إلى أي مكان.
 */
(function (root) {
  'use strict';
  const DB_NAME = 'mu3abbi-istibyanat';
  const DB_VER = 1;
  const STORE = 'profiles';
  let dbPromise = null;

  function open() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        if (!root.indexedDB) { reject(new Error('هذا المتصفح لا يدعم الحفظ الداخلي. استعمل كروم أو إيدج أو فايرفوكس حديثًا.')); return; }
        const req = root.indexedDB.open(DB_NAME, DB_VER);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error || new Error('تعذّر فتح قاعدة بيانات المتصفح.'));
        req.onblocked = () => reject(new Error('أغلق نوافذ البرنامج الأخرى المفتوحة ثم أعد التحميل.'));
      });
    }
    return dbPromise;
  }

  function wrap(req) {
    return new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function store(mode) {
    const db = await open();
    return db.transaction(STORE, mode).objectStore(STORE);
  }

  const Store = {
    async all() {
      const list = await wrap((await store('readonly')).getAll());
      return list.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    },
    async get(id) { return wrap((await store('readonly')).get(id)); },
    async put(profile) {
      profile.updatedAt = Date.now();
      await wrap((await store('readwrite')).put(profile));
      Store.persist();
      return profile;
    },
    async remove(id) { return wrap((await store('readwrite')).delete(id)); },

    /** يطلب من المتصفح ألا يمسح البيانات تلقائيًا عند امتلاء القرص */
    async persist() {
      try {
        if (navigator.storage && navigator.storage.persist && !(await navigator.storage.persisted())) {
          await navigator.storage.persist();
        }
      } catch (e) { /* لا بأس */ }
    },
    async usage() {
      try {
        if (navigator.storage && navigator.storage.estimate) {
          const e = await navigator.storage.estimate();
          return e.usage || 0;
        }
      } catch (e) { /* لا بأس */ }
      return null;
    },
  };

  root.Store = Store;
})(self);
