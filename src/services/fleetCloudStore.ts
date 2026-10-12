import { doc, getDoc, setDoc } from 'firebase/firestore';
import { firebaseReady, firestore, isFirebaseConfigured } from './firebase';

const fleetDocument = firestore ? doc(firestore, 'fleet', 'company') : null;
const routeDocument = firestore ? doc(firestore, 'fleet', 'routes') : null;
let fleetSaveQueue: Promise<void> = Promise.resolve();
let routeSaveQueue: Promise<void> = Promise.resolve();

export const readCloudFleetData = async <T>(): Promise<T | null> => {
  if (!isFirebaseConfigured || !fleetDocument) return null;
  try {
    await firebaseReady;
    const snapshot = await getDoc(fleetDocument);
    return snapshot.exists() ? snapshot.data() as T : null;
  } catch (error) {
    console.error('Could not read fleet data from Firestore', error);
    return null;
  }
};

export const saveCloudFleetDataStrict = <T extends object>(data: T): Promise<void> => {
  if (!isFirebaseConfigured || !fleetDocument) return;
  const save = fleetSaveQueue.then(async () => {
    await firebaseReady;
    await setDoc(fleetDocument, data);
  });
  fleetSaveQueue = save.catch(() => undefined);
  return save;
};

export const saveCloudFleetData = async <T extends object>(data: T): Promise<void> => {
  try {
    await saveCloudFleetDataStrict(data);
  } catch (error) {
    console.error('Could not save fleet data to Firestore', error);
  }
};

export const readCloudRouteEntries = async <T>(): Promise<T[] | null> => {
  if (!isFirebaseConfigured || !routeDocument) return null;
  await firebaseReady;
  const snapshot = await getDoc(routeDocument);
  const records = snapshot.exists() ? snapshot.data().records : null;
  return Array.isArray(records) ? records as T[] : null;
};

export const saveCloudRouteEntries = async <T extends object>(records: T[]): Promise<void> => {
  if (!isFirebaseConfigured || !routeDocument) return;
  const save = routeSaveQueue.then(async () => {
    await firebaseReady;
    await setDoc(routeDocument, { records });
  });
  routeSaveQueue = save.catch(() => undefined);
  await save;
};
