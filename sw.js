// Service Worker HoopTrack — reçoit les notifications push même app fermée, et gère le clic
// dessus (ouvre/ramène l'app au premier plan). Volontairement minimal : ce fichier ne gère QUE
// les notifications, il ne met rien en cache (pas de mode hors-ligne pour l'instant).

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("push", (event) => {
  if (!event.data) return;
  let payload;
  try {
    payload = event.data.json();
  } catch (e) {
    payload = { title: "HoopTrack", body: event.data.text() };
  }
  const title = payload.title || "HoopTrack";
  const options = {
    body: payload.body || "",
    icon: "/icon-512.png",
    badge: "/icon-512.png",
    data: { url: payload.url || "/" },
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

// Au clic sur la notification : ramène un onglet déjà ouvert au premier plan s'il y en a un,
// sinon en ouvre un nouveau.
self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = event.notification.data?.url || "/";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ("focus" in client) return client.focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow(targetUrl);
    })
  );
});
