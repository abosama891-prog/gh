let registrationPromise: Promise<ServiceWorkerRegistration> | undefined;
let initialized = false;

const registerServiceWorker = () => {
  registrationPromise ??= navigator.serviceWorker.register('/sw.js', { updateViaCache: 'none' });
  return registrationPromise;
};

const announceWaitingUpdate = (registration: ServiceWorkerRegistration) => {
  if (navigator.serviceWorker.controller && registration.waiting) {
    window.dispatchEvent(new CustomEvent('sw-update-available'));
  }
};

export const initializeServiceWorker = () => {
  if (!('serviceWorker' in navigator) || initialized) return;
  initialized = true;

  let hadController = Boolean(navigator.serviceWorker.controller);
  let refreshing = false;

  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hadController) {
      hadController = true;
      return;
    }
    if (refreshing) return;
    refreshing = true;
    window.location.reload();
  });

  const checkForUpdate = () => {
    void registerServiceWorker()
      .then((registration) => {
        registration.addEventListener('updatefound', () => {
          const newWorker = registration.installing;
          if (!newWorker) return;
          const announceInstalledUpdate = () => {
            if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
              window.dispatchEvent(new CustomEvent('sw-update-available'));
            }
          };
          newWorker.addEventListener('statechange', announceInstalledUpdate);
          announceInstalledUpdate();
        });

        announceWaitingUpdate(registration);
        return registration.update().then(() => announceWaitingUpdate(registration));
      })
      .catch((error: unknown) => console.error('Service worker registration or update failed', error));
  };

  window.addEventListener('load', checkForUpdate, { once: true });
  window.addEventListener('pageshow', checkForUpdate);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') checkForUpdate();
  });
  window.setInterval(checkForUpdate, 60 * 60 * 1000);
};

export const applyUpdate = async (): Promise<void> => {
  if (!('serviceWorker' in navigator)) {
    window.location.reload();
    return;
  }

  const registration = await registerServiceWorker();
  await registration.update();
  if (registration?.waiting) {
    registration.waiting.postMessage('SKIP_WAITING');
    return;
  }

  window.location.reload();
};
