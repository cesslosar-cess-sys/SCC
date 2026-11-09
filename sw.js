const CACHE_NAME = 'taller-cess-gold-v18';
const STATIC_ASSETS = [
    './icon.svg',
    './manifest.json'
];

// Instalación inmediata y reemplazo de cualquier SW anterior
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[SW v18] Instalando versión CESS Gold...');
            return cache.addAll(STATIC_ASSETS);
        })
    );
});

// Activación: Purga forzada e inmediata de todas las cachés viejas
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(
                keys.map((key) => {
                    if (key !== CACHE_NAME) {
                        console.log('[SW v18] Eliminando caché obsoleta:', key);
                        return caches.delete(key);
                    }
                })
            );
        }).then(() => {
            return self.clients.claim();
        })
    );
});

// Interceptor Fetch: NETWORK-FIRST para HTML y navegación para ver cambios al instante
self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') return;

    const url = new URL(event.request.url);

    // 1. REGLA ESTRICTA DE TIEMPO REAL: Las peticiones a la API o Streams NUNCA se cachean
    if (url.pathname.startsWith('/api/') || url.pathname.includes('/api/')) {
        event.respondWith(
            fetch(event.request, { cache: 'no-store' }).catch(() => {
                return new Response(JSON.stringify({ error: 'offline', offline: true, data: [] }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            })
        );
        return;
    }

    const isNavigationOrHtml = event.request.mode === 'navigate' || 
                               url.pathname === '/' || 
                               url.pathname.endsWith('.html') ||
                               url.pathname.endsWith('server.js');

    if (isNavigationOrHtml) {
        // Para HTML: Buscar SIEMPRE en la red primero para garantizar que los cambios se vean de inmediato
        event.respondWith(
            fetch(event.request)
                .then((networkResponse) => {
                    if (networkResponse && networkResponse.status === 200) {
                        const clone = networkResponse.clone();
                        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
                    }
                    return networkResponse;
                })
                .catch(() => {
                    // Solo si no hay internet (offline en patio), usamos la copia guardada
                    return caches.match(event.request).then(cached => cached || caches.match('./index.html'));
                })
        );
        return;
    }

    // Para otros recursos (iconos, fuentes, CSS estáticos)
    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
                // Actualizar en fondo
                fetch(event.request).then((netRes) => {
                    if (netRes && netRes.status === 200) {
                        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, netRes));
                    }
                }).catch(() => {});
                return cachedResponse;
            }

            return fetch(event.request).then((networkResponse) => {
                if (networkResponse && (networkResponse.status === 200 || networkResponse.type === 'opaque')) {
                    const clone = networkResponse.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
                }
                return networkResponse;
            });
        })
    );
});
