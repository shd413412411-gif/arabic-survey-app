// ارفع رقم الإصدار عند كل نشر جديد لتحديث ملفات التطبيق المخزنة.
const CACHE_NAME = 'survey-app-v1';
const APP_SHELL = ['/', '/index.html', '/manifest.json'];
const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(APP_SHELL);

    // تخزين ملفات JavaScript وCSS التي يشير إليها index.html لتشغيل الواجهة دون اتصال.
    try {
      const htmlResponse = await fetch('/index.html', { cache: 'reload' });
      const html = await htmlResponse.text();
      const assetPaths = [...html.matchAll(/(?:src|href)=["']([^"']+\.(?:js|css)(?:\?[^"']*)?)["']/g)]
        .map((match) => match[1])
        .filter((path) => path.startsWith('/') || !path.startsWith('http'));
      await Promise.all(assetPaths.map(async (path) => {
        try {
          const assetUrl = new URL(path, self.location.origin).href;
          const response = await fetch(assetUrl);
          if (response.ok) await cache.put(assetUrl, response);
        } catch (error) {
          // يستمر التثبيت حتى لو تعذر تخزين مورد ثانوي.
        }
      }));
    } catch (error) {
      // بعض بيئات المعاينة لا تعرض index.html مباشرة؛ يخزن الجلب الاعتيادي الموارد لاحقًا.
    }

    // تخزين SheetJS اختياري؛ إن تعذر جلبه لا يتعطل تثبيت التطبيق أو الإدخال المحلي.
    try {
      const response = await fetch(XLSX_URL, { mode: 'cors' });
      if (response.ok) await cache.put(XLSX_URL, response);
    } catch (error) {
      // تتطلب وظائف Excel اتصالًا حتى تُحمّل المكتبة لأول مرة.
    }

    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys
      .filter((key) => key.startsWith('survey-app-') && key !== CACHE_NAME)
      .map((key) => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  if (request.url === XLSX_URL) {
    event.respondWith((async () => {
      const cached = await caches.match(XLSX_URL);
      if (cached) return cached;
      try {
        const response = await fetch(request);
        if (response.ok) {
          const cache = await caches.open(CACHE_NAME);
          cache.put(XLSX_URL, response.clone());
        }
        return response;
      } catch (error) {
        return new Response('SheetJS is not available offline yet.', {
          status: 503,
          headers: { 'Content-Type': 'text/plain; charset=utf-8' },
        });
      }
    })());
    return;
  }

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    const cached = await caches.match(request);
    if (cached) return cached;
    try {
      const response = await fetch(request);
      if (response.ok && (request.destination === 'script' || request.destination === 'style' || request.destination === 'document' || request.destination === 'font')) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(request, response.clone());
      }
      return response;
    } catch (error) {
      return (await caches.match('/index.html')) || new Response('الموقع غير متاح دون اتصال بعد.', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
    }
  })());
});
