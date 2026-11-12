/*
 * مُعبِّئ الاستبيانات — عامل الخدمة
 * يحفظ ملفات البرنامج نفسها (لا بيانات المستخدم) ليعمل البرنامج دون إنترنت بعد أول فتح.
 * يجلب النسخة الأحدث من الموقع دائمًا إن توفر الاتصال، ويرجع إلى المحفوظ عند انقطاعه.
 * عند تعديل ملفات البرنامج غيّر رقم الإصدار في السطر التالي.
 */
const CACHE = 'mu3abbi-v1.1';
const FILES = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'engine.js',
  'storage.js',
  'templates.js',
  'lib/pizzip.min.js',
  'lib/docxtemplater.js',
  'fonts/qomra-medium.woff2',
  'assets/logo.png',
  'assets/favicon.ico',
  'assets/favicon-32.png',
  'assets/icon-180.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES).catch(() => {})).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('mu3abbi-') && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(
    fetch(req)
      .then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
