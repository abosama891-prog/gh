import bcrypt from 'bcryptjs';
import { addDoc, collection, doc, getDoc, onSnapshot, runTransaction } from 'firebase/firestore';
import { firebaseReady, firestore } from './firebase';

export type FleetRole = 'admin' | 'viewer';

export type FleetAdmin = {
  username: string;
  role: 'admin';
};

export type FleetGuest = {
  id: string;
  passwordHash: string;
  employeeName: string;
  role: 'viewer';
  active: boolean;
  isGuest: true;
  createdAt: string;
};

export type FleetGuestInput = { password: string; employeeName: string };

export type FleetUser = {
  id: string;
  username: string;
  email?: string;
  employeeName?: string;
  password: string;
  role: FleetRole;
  active: boolean;
  isGuest?: boolean;
  guestName?: string;
  createdAt: string;
};

const ADMINS_KEY = 'cars-site-admins-v1';
const ADMIN_PASSWORD_HASHES_KEY = 'cars-site-admin-password-hashes-v1';
const USERS_DOCUMENT = firestore ? doc(firestore, 'fleet', 'users') : null;
const LOGIN_LOGS_COLLECTION = firestore ? collection(firestore, 'fleet', 'login-logs', 'items') : null;
const defaultAdmins: FleetAdmin[] = [
  { username: 'admin', role: 'admin' },
  { username: 'gh', role: 'admin' },
];

export const getAdminAuthEmail = (username: string): string => `${username.trim().toLowerCase()}@aboghaly.local`;

const readAdminPasswordHashes = (): Record<string, string> => {
  try {
    const stored = localStorage.getItem(ADMIN_PASSWORD_HASHES_KEY);
    if (!stored) return {};
    const parsed = JSON.parse(stored) as unknown;
    if (!parsed || typeof parsed !== 'object') return {};
    return Object.fromEntries(Object.entries(parsed as Record<string, unknown>).filter(([, value]) => typeof value === 'string')) as Record<string, string>;
  } catch (error) {
    console.error('Could not read stored admin password hashes', error);
    return {};
  }
};

const saveAdminPasswordHashes = (hashes: Record<string, string>): void => {
  try {
    localStorage.setItem(ADMIN_PASSWORD_HASHES_KEY, JSON.stringify(hashes));
  } catch (error) {
    console.error('Could not save admin password hashes', error);
  }
};

export const saveAdminPasswordHash = (username: string, password: string): void => {
  const normalized = username.trim().toLowerCase();
  const hashes = readAdminPasswordHashes();
  hashes[normalized] = bcrypt.hashSync(password, 10);
  saveAdminPasswordHashes(hashes);
};

export const verifyAdminPassword = async (username: string, password: string): Promise<boolean> => {
  const normalized = username.trim().toLowerCase();
  const hashes = readAdminPasswordHashes();
  const storedHash = hashes[normalized];
  if (!storedHash) return false;
  if (/^\$2[aby]\$/.test(storedHash)) return bcrypt.compare(password, storedHash);
  return password === storedHash;
};

export const verifyGuestPassword = async (input: string, stored: string): Promise<boolean> => {
  if (/^\$2[aby]\$/.test(stored)) return bcrypt.compare(input, stored);
  return input === stored;
};

export const recordLoginAttempt = async (username: string, success: boolean): Promise<void> => {
  if (!LOGIN_LOGS_COLLECTION) throw new Error('Firestore login log collection is not configured');
  await firebaseReady;
  await addDoc(LOGIN_LOGS_COLLECTION, {
    timestamp: new Date().toISOString(),
    username,
    success,
    userAgent: typeof navigator === 'undefined' ? '' : navigator.userAgent,
  });
};

const parseGuests = (users: unknown): FleetGuest[] => {
  if (!Array.isArray(users)) return [];
  return users.flatMap((item): FleetGuest[] => {
    if (typeof item !== 'object' || item === null) return [];
    const user = item as Record<string, unknown>;
    const employeeName = typeof user.employeeName === 'string'
      ? user.employeeName.trim()
      : typeof user.guestName === 'string'
        ? user.guestName.trim()
        : typeof user.username === 'string'
          ? user.username.trim()
          : '';
    if (!employeeName) return [];
    return [{
      id: typeof user.id === 'string' ? user.id : `guest-${employeeName.toLowerCase()}`,
      passwordHash: typeof user.passwordHash === 'string' && user.passwordHash
        ? user.passwordHash
        : typeof user.password === 'string'
          ? user.password
          : '',
      employeeName,
      role: 'viewer',
      active: typeof user.active === 'boolean' ? user.active : true,
      isGuest: true,
      createdAt: typeof user.createdAt === 'string' ? user.createdAt : new Date().toISOString(),
    }];
  });
};

export const readAdmins = (): FleetAdmin[] => {
  try {
    const stored = localStorage.getItem(ADMINS_KEY);
    if (!stored) {
      localStorage.setItem(ADMINS_KEY, JSON.stringify(defaultAdmins));
      return [...defaultAdmins];
    }
    const admins = JSON.parse(stored) as FleetAdmin[];
    return Array.isArray(admins) && admins.length ? admins : [...defaultAdmins];
  } catch {
    return [...defaultAdmins];
  }
};

export const saveAdmins = (admins: FleetAdmin[]): void => {
  localStorage.setItem(ADMINS_KEY, JSON.stringify(admins));
};

export const getCloudFleetUsers = async (): Promise<FleetGuest[]> => {
  if (!USERS_DOCUMENT) throw new Error('Firestore is not configured');
  await firebaseReady;
  const snapshot = await getDoc(USERS_DOCUMENT);
  if (!snapshot.exists()) return [];
  return parseGuests(snapshot.data().users);
};

export const subscribeToFleetUsers = (
  callback: (guests: FleetGuest[]) => void,
  onError?: (error: Error) => void,
): (() => void) => {
  if (!USERS_DOCUMENT) {
    const error = new Error('Firestore is not configured');
    console.error('Failed to subscribe to guests', error);
    onError?.(error);
    return () => {};
  }
  return onSnapshot(USERS_DOCUMENT, (snapshot) => {
    callback(snapshot.exists() ? parseGuests(snapshot.data().users) : []);
  }, (error) => {
    console.error('Failed to subscribe to guests', error);
    onError?.(error);
  });
};

export const addGuest = async ({ password, employeeName }: FleetGuestInput): Promise<FleetGuest> => {
  if (!USERS_DOCUMENT) throw new Error('Firestore is not configured');
  const normalizedEmployeeName = employeeName.trim();
  if (!password.trim()) throw new Error('أدخل كلمة المرور');
  if (!normalizedEmployeeName) throw new Error('أدخل اسم الموظف');
  if (new TextEncoder().encode(password).length > 72) throw new Error('يجب ألا تتجاوز كلمة المرور 72 بايت');
  const passwordHash = await bcrypt.hash(password, 10);
  const guest: FleetGuest = {
    id: `guest-${crypto.randomUUID()}`,
    passwordHash,
    employeeName: normalizedEmployeeName,
    role: 'viewer',
    active: true,
    isGuest: true,
    createdAt: new Date().toISOString(),
  };
  await firebaseReady;
  await runTransaction(firestore!, async (transaction) => {
    const snapshot = await transaction.get(USERS_DOCUMENT);
    const guests = snapshot.exists() ? parseGuests(snapshot.data().users) : [];
    const duplicate = await Promise.all(guests
      .filter((existingGuest) => existingGuest.passwordHash)
      .map((existingGuest) => bcrypt.compare(password, existingGuest.passwordHash)));
    if (duplicate.some(Boolean)) {
      throw new Error('كلمة المرور مستخدمة بالفعل');
    }
    transaction.set(USERS_DOCUMENT, { users: [...guests, guest] });
  });
  return guest;
};

export const updateGuest = async (guestId: string, { password, employeeName }: FleetGuestInput): Promise<void> => {
  if (!USERS_DOCUMENT) throw new Error('Firestore is not configured');
  const normalizedEmployeeName = employeeName.trim();
  if (!normalizedEmployeeName) throw new Error('أدخل اسم الموظف');
  if (password && new TextEncoder().encode(password).length > 72) throw new Error('يجب ألا تتجاوز كلمة المرور 72 بايت');
  const passwordHash = password ? await bcrypt.hash(password, 10) : '';
  await firebaseReady;
  await runTransaction(firestore!, async (transaction) => {
    const snapshot = await transaction.get(USERS_DOCUMENT);
    if (!snapshot.exists()) return;
    const guests = parseGuests(snapshot.data().users);
    const duplicate = await Promise.all(guests
      .filter((guest) => guest.id !== guestId && guest.passwordHash)
      .map((guest) => bcrypt.compare(password, guest.passwordHash)));
    if (password && duplicate.some(Boolean)) {
      throw new Error('كلمة المرور مستخدمة بالفعل');
    }
    const nextGuests = guests.map((guest) => guest.id === guestId
      ? { ...guest, passwordHash: passwordHash || guest.passwordHash, employeeName: normalizedEmployeeName }
      : guest);
    transaction.set(USERS_DOCUMENT, { users: nextGuests });
  });
};

export const setGuestActive = async (guestId: string, active: boolean): Promise<void> => {
  if (!USERS_DOCUMENT) throw new Error('Firestore is not configured');
  await firebaseReady;
  await runTransaction(firestore!, async (transaction) => {
    const snapshot = await transaction.get(USERS_DOCUMENT);
    if (!snapshot.exists()) return;
    const guests = parseGuests(snapshot.data().users);
    transaction.set(USERS_DOCUMENT, {
      users: guests.map((guest) => guest.id === guestId ? { ...guest, active } : guest),
    });
  });
};

export const deleteGuest = async (guestId: string): Promise<void> => {
  if (!USERS_DOCUMENT) throw new Error('Firestore is not configured');
  await firebaseReady;
  await runTransaction(firestore!, async (transaction) => {
    const snapshot = await transaction.get(USERS_DOCUMENT);
    if (!snapshot.exists()) return;
    const guests = parseGuests(snapshot.data().users);
    transaction.set(USERS_DOCUMENT, { users: guests.filter((guest) => guest.id !== guestId) });
  });
};
