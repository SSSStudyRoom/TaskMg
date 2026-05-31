self.addEventListener('install', (e) => {
  console.log('[Service Worker] Install');
});
self.addEventListener('fetch', (e) => {
  // 基本はネットワーク通信を優先（常に最新のデータを取る）
});
