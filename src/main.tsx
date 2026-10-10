import React from 'react';
import ReactDOM from 'react-dom/client';
import CarsApp from './CarsApp';
import './styles/globals.css';

if ('serviceWorker' in navigator) {
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
}

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(
  <React.StrictMode>
    <CarsApp />
  </React.StrictMode>
);