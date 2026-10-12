// This file contains Vite environment definitions. 

/// <reference types="vite/client" />

interface BeforeInstallPromptEvent extends Event {
	prompt: () => Promise<void>;
	userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }>;
}
