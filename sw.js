const CACHE_NAME = 'inventario-pro-v7.4-cache-v1';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './styles.css',
  './script.js',
  './manifest.json',
  // Asegúrate de tener tu logo en la carpeta raíz, si no tienes la imagen,
  // el service worker podría dar error al intentar cachearla inicialmente.
  // Si usas otro nombre de imagen, actualízalo aquí.
  './logo.png' 
];

// Evento de Instalación: Cacha los archivos estáticos críticos
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[Service Worker] Caching app shell');
        return cache.addAll(ASSETS_TO_CACHE);
      })
      .then(() => self.skipWaiting()) // Activar el SW inmediatamente
  );
});

// Evento de Activación: Limpia cachés antiguas si actualizas la versión
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[Service Worker] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim()) // Tomar control de los clientes inmediatamente
  );
});

// Evento Fetch: Estrategia "Cache First, falling back to Network"
// Esto permite que las librerías CDN (Tailwind, etc.) se guarden localmente tras la primera carga
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        // 1. Si está en caché, lo devuelve
        if (response) {
          return response;
        }

        // 2. Si no, hace la petición a la red
        return fetch(event.request).then(
          (networkResponse) => {
            // Verificamos si la respuesta es válida
            if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
              return networkResponse;
            }

            // Clonamos la respuesta para guardarla en caché (el stream solo se lee una vez)
            const responseToCache = networkResponse.clone();

            caches.open(CACHE_NAME)
              .then((cache) => {
                // Guardamos dinámicamente cualquier recurso nuevo (como las fuentes y scripts CDN)
                cache.put(event.request, responseToCache);
              });

            return networkResponse;
          }
        );
      })
      .catch(() => {
        // Fallback opcional para cuando no hay red y no está en caché
        console.log('[Service Worker] Fetch failed; returning offline page instead.', event.request.url);
      })
  );
});