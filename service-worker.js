const CACHE_VERSION = 'app-v1.0.0';
const CACHE_STATIC = `${CACHE_VERSION}-static`;

const staticAssets = [
  '/',
  '/index.html',
  '/english.html',
  '/chinese.html',
  '/json/english.json',
  '/json/chinese.json'
];

// 安裝事件 - 快取靜態資源
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_STATIC).then((cache) => {
      // 逐個添加，某個失敗不影響其他
      return Promise.all(
        staticAssets.map((url) => {
          return cache.add(url).catch((error) => {
            console.log(`[SW] 快取失敗 [${url}]:`, error.message);
          });
        })
      );
    })
  );
  // self.skipWaiting();  // 註解掉：讓新 sw 進入 waiting，等待用戶確認更新
});

// 啟用事件 - 清除舊版本快取
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          // 刪除不匹配當前版本的舊快取
          if (cacheName !== CACHE_STATIC) {
            console.log('刪除舊快取:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
  // 通知所有客戶端更新完成
  self.clients.matchAll().then((clients) => {
    clients.forEach((client) => {
      client.postMessage({ type: 'SW_UPDATED', version: CACHE_VERSION });
    });
  });
});

// Fetch 事件 - 實現快取策略
self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // 跳過非同源和特殊協議
  if (!url.origin.includes(self.location.origin)) {
    return;
  }

  // 跳過非 GET 請求
  if (request.method !== 'GET') {
    return;
  }

  // JSON 數據和 API 請求 - 快取優先，失敗才用網路
  if (url.pathname.includes('/api/') || url.pathname.endsWith('.json')) {
    event.respondWith(
      caches.match(request).then((cached) => {
        if (cached) {
          console.log(`[SW] ✅ JSON 快取命中: ${url.pathname}`);
          return cached;
        }
        // 快取無，才從網路獲取
        console.log(`[SW] 🌐 JSON 無快取，從網路獲取: ${url.pathname}`);
        return fetch(request)
          .then((response) => {
            if (response.ok) {
              console.log(`[SW] 💾 JSON 網路獲取成功，快取存儲: ${url.pathname}`);
              caches.open(CACHE_STATIC).then((cache) => {
                cache.put(request, response.clone());
              });
            }
            return response;
          })
          .catch((error) => {
            // 網路也失敗，返回空陣列
            console.log(`[SW] ❌ JSON 網路失敗，返回空陣列: ${url.pathname}`);
            return new Response(JSON.stringify([]), {
              headers: { 'Content-Type': 'application/json' },
              status: 200
            });
          });
      })
    );
    return;
  }

  // HTML 頁面 - 整頁快取（快取優先）
  if (request.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname === '/') {
    event.respondWith(
      caches.open(CACHE_STATIC).then((cache) => {
        // 規範化 URL 路徑用於快取查詢
        let pathToMatch = url.pathname;
        // 如果是根路径且不是以 / 結尾，補上 / 以便匹配
        if (pathToMatch === '') {
          pathToMatch = '/';
        }
        // 如果是根路径，也嘗試匹配 /index.html
        const urlsToTry = pathToMatch === '/'
          ? ['/', '/index.html']
          : [pathToMatch];

        console.log('[SW] 導航請求:', pathToMatch, '嘗試快取鍵:', urlsToTry);

        // 先嘗試精確匹配
        return cache.match(request).then(async (cachedResponse) => {
          if (cachedResponse) {
            console.log('[SW] 精確匹配快取成功:', pathToMatch);
            return cachedResponse;
          }

          // 精確匹配失敗，嘗試規範化的 URL 列表
          for (const urlToTry of urlsToTry) {
            try {
              const cached = await cache.match(urlToTry);
              if (cached) {
                console.log('[SW] 規範化路徑匹配成功:', urlToTry);
                return cached;
              }
            } catch (e) {
              console.log('[SW] 規範化路徑匹配失敗:', urlToTry, e.message);
            }
          }

          // 快取中沒有，嘗試網路請求
          console.log('[SW] 快取未命中，嘗試網路請求:', pathToMatch);
          return fetch(request).then((response) => {
            if (response.ok) {
              console.log('[SW] 網路請求成功，快取頁面:', pathToMatch);
              cache.put(request, response.clone());
            }
            return response;
          }).catch((error) => {
            console.log('[SW] 網路請求失敗，進入 Fallback:', pathToMatch, error.message);

            const fallbacks = [
              '/index.html',
              url.pathname,
              '/'
            ];

            // 異步處理 fallback（避免Promise直接作為真值檢查）
            return (async () => {
              for (const fallback of fallbacks) {
                try {
                  console.log('[SW] 嘗試從快取獲取:', fallback);
                  const cached = await cache.match(fallback);
                  if (cached) {
                    console.log('[SW] 快取命中:', fallback);
                    return cached;
                  }
                } catch (e) {
                  console.log('[SW] 快取查詢失敗:', fallback, e.message);
                }
              }

              // 都失敗，返回任何可用的 HTML
              console.log('[SW] Fallback 列表都失敗，嘗試查找任何 HTML');
              return cache.keys().then((keys) => {
                const htmlKey = keys.find(k => k.url.endsWith('.html'));
                if (htmlKey) {
                  console.log('[SW] 找到快取 HTML:', htmlKey.url);
                  return cache.match(htmlKey);
                }

                // 最後的 fallback
                console.log('[SW] 無快取可用，返回離線提示');
                return new Response(
                  '<h1>離線模式</h1><p>無可用內容。請檢查網路連接。</p>',
                  { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
                );
              });
            })();
          });
        });
      })
    );
    return;
  }

  // 靜態資源 - 快取優先，失敗才用網路
  event.respondWith(
    caches.match(request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }

      return fetch(request)
        .then((response) => {
          if (response && response.status === 200) {
            cache.put(request, response.clone());
          }
          return response;
        })
        .catch(() => {
          // 快取無且網路失敗
          return new Response('', { status: 204 });
        });
    })
  );
});

// 接收來自主線程的消息
self.addEventListener('message', (event) => {
  console.log('[SW] 收到訊息:', event.data.type);
  if (event.data.type === 'SKIP_WAITING') {
    console.log('[SW] 執行 skipWaiting()');
    self.skipWaiting();
  }
});

// 全域錯誤處理
self.addEventListener('error', (event) => {
  console.error('[SW] 錯誤:', event.error);
});
