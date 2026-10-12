import {
  addDoc,
  collection,
  limit,
  onSnapshot,
  orderBy,
  query,
  type Unsubscribe,
} from 'firebase/firestore';
import { firebaseAuth, firebaseReady, firestore } from './firebase';

export type GpsLog = {
  id: string;
  timestamp: string;
  latitude: number;
  longitude: number;
  accuracy?: number;
  vehicleId?: string;
  driver: string;
  note?: string;
};

const gpsLogsCollection = firestore ? collection(firestore, 'fleet', 'gps-logs', 'items') : null;

export const saveGpsLog = async (log: Omit<GpsLog, 'id'>): Promise<GpsLog> => {
  await firebaseReady;
  if (!firebaseAuth?.currentUser) {
    throw new Error('يجب تسجيل الدخول إلى Firebase قبل حفظ سجل GPS');
  }
  if (!gpsLogsCollection) {
    throw new Error('خدمة Firestore غير مهيأة لحفظ سجلات GPS');
  }

  const timestamp = log.timestamp || new Date().toISOString();
  const document = await addDoc(gpsLogsCollection, {
    ...log,
    timestamp,
  });
  return { ...log, timestamp, id: document.id };
};

export const subscribeToGpsLogs = (
  onLogs: (logs: GpsLog[]) => void,
  onError: (error: Error) => void,
): Unsubscribe => {
  if (!firebaseAuth?.currentUser) {
    onError(new Error('يجب تسجيل الدخول إلى Firebase لمزامنة سجلات GPS'));
    return () => undefined;
  }
  if (!gpsLogsCollection) {
    onError(new Error('خدمة Firestore غير مهيأة لمزامنة سجلات GPS'));
    return () => undefined;
  }

  const logsQuery = query(gpsLogsCollection, orderBy('timestamp', 'desc'), limit(500));
  return onSnapshot(
    logsQuery,
    (snapshot) => {
      onLogs(snapshot.docs.map((document) => ({ ...document.data(), id: document.id }) as GpsLog));
    },
    (error) => {
      console.error('Failed to subscribe to GPS logs', error);
      onError(error);
    },
  );
};
