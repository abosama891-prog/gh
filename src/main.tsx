import React from 'react';
import ReactDOM from 'react-dom/client';
import CarsApp from './CarsApp';
import './styles/globals.css';
import { initializePwaInstall } from './services/pwaInstall';

initializePwaInstall();

if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  let refreshing = false;
  let hasActiveController = Boolean(navigator.serviceWorker.controller);

  const checkForServiceWorkerUpdate = () => {
    void navigator.serviceWorker.register('/sw.js', { updateViaCache: 'none' })
      .then((registration) => registration.update())
      .catch((error: unknown) => console.error('Service worker registration or update failed', error));
  };

  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hasActiveController) {
      hasActiveController = true;
      return;
    }
    if (refreshing) return;
    refreshing = true;
    window.location.reload();
  });

  navigator.serviceWorker.addEventListener('message', (event: MessageEvent<unknown>) => {
    if (typeof event.data === 'object' && event.data !== null && 'type' in event.data && event.data.type === 'SW_UPDATED') {
      console.info('تم تحديث التطبيق');
    }
  });

  const startUpdateChecks = () => {
    checkForServiceWorkerUpdate();
    window.setInterval(checkForServiceWorkerUpdate, 5 * 60 * 1000);
    window.addEventListener('pageshow', checkForServiceWorkerUpdate);
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') checkForServiceWorkerUpdate();
    });
  };

  if (document.readyState === 'complete') startUpdateChecks();
  else window.addEventListener('load', startUpdateChecks, { once: true });
} else if (import.meta.env.DEV && 'serviceWorker' in navigator) {
  void navigator.serviceWorker.getRegistrations().then((registrations) => {
    const appRegistrations = registrations.filter((registration) => (
      [registration.active, registration.waiting, registration.installing]
        .some((worker) => worker?.scriptURL === new URL('/sw.js', window.location.origin).href)
    ));
    return Promise.all(appRegistrations.map((registration) => registration.unregister()));
  }).then(() => caches.keys()).then((cacheNames) => Promise.all(
    cacheNames
      .filter((cacheName) => cacheName.startsWith('aboghaly-cars-v'))
      .map((cacheName) => caches.delete(cacheName)),
  )).catch((error: unknown) => console.error('Failed to clear the app service worker cache in development', error));
}

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(
  <React.StrictMode>
    <CarsApp />
  </React.StrictMode>
);