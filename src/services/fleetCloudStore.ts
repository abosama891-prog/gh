import { doc, getDoc, setDoc } from 'firebase/firestore';
import { firebaseReady, firestore, isFirebaseConfigured } from './firebase';

const fleetDocument = firestore ? doc(firestore, 'fleet', 'company') : null;

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

export const saveCloudFleetData = async <T extends object>(data: T): Promise<void> => {
  if (!isFirebaseConfigured || !fleetDocument) return;
  try {
    await firebaseReady;
    await setDoc(fleetDocument, data);
  } catch (error) {
    console.error('Could not save fleet data to Firestore', error);
  }
};
