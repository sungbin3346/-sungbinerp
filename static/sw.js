// =============================================================================
// Service Worker — 영업 관리 ERP PWA
// 오프라인 읽기 전용 캐싱 전략
// =============================================================================
const CACHE_NAME = 'erp-cache-v1';

// 캐시할 리소스 목록 (오프라인 지원)
const CACHE_URLS = [
  '/',
  '/app/static/manifest.json',
];

// 설치 이벤트: 핵심 리소스 사전 캐싱
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(CACHE_URLS).catch(() => {
        // 일부 URL 캐싱 실패해도 계속 진행
        console.log('SW: 일부 리소스 캐싱 실패 (무시)');
      });
    })
  );
  self.skipWaiting();
});

// 활성화 이벤트: 이전 캐시 삭제
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// Fetch 이벤트: 네트워크 우선, 실패 시 캐시 반환
self.addEventListener('fetch', (event) => {
  // POST/non-GET 요청은 캐시하지 않음
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // 성공한 응답은 캐시에 저장
        if (response.status === 200) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }
        return response;
      })
      .catch(() => {
        // 네트워크 실패 시 캐시에서 반환
        return caches.match(event.request);
      })
  );
});
