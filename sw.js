const CACHE_NAME = 'Smart-control-beton-v2';

// 1. Lister précisément le fichier HTML et ses dépendances
const urlsToCache = [
  './',
  './index.html', // ou le nom de votre fichier HTML autonome
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Installation du Service Worker et mise en cache
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(urlsToCache);
    })
  );
  self.skipWaiting();
});

// Activation et nettoyage des anciens caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Gestion des requêtes : Priorité au Cache (Cache-First)
self.addEventListener('fetch', (event) => {
  // Ignorer la communication Streamlit en ligne
  if (event.request.url.includes('_stcore') || event.request.url.includes('stream')) {
    return;
  }

  // Chercher d'abord dans le cache local, puis basculer sur le réseau
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse; // Retourne immédiatement la version hors-ligne
      }
      return fetch(event.request);
    })
  );
});
