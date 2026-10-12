let deferredPrompt: BeforeInstallPromptEvent | null = null;
let initialized = false;

export const initializePwaInstall = () => {
  if (initialized) return;
  initialized = true;

  window.addEventListener('beforeinstallprompt', (event: BeforeInstallPromptEvent) => {
    event.preventDefault();
    deferredPrompt = event;
    window.dispatchEvent(new Event('pwa-install-available'));
  });

  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    window.dispatchEvent(new Event('pwa-installed'));
  });
};

export const getDeferredPrompt = () => deferredPrompt;

export const clearDeferredPrompt = () => {
  deferredPrompt = null;
};
