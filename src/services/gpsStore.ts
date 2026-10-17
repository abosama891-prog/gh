import {
  addDoc,
  collection,
  getDocs,
  limit,
  onSnapshot,
  orderBy,
  query,
  startAfter,
  type DocumentData,
  type QueryDocumentSnapshot,
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

export type GpsLogsPage = {
  logs: GpsLog[];
  cursor: QueryDocumentSnapshot<DocumentData> | null;
  hasMore: boolean;
};

export type GpsLogCursor = QueryDocumentSnapshot<DocumentData>;

const gpsLogsCollection = firestore ? collection(firestore, 'fleet', 'gps-logs', 'items') : null;
const GPS_LOGS_PAGE_SIZE = 50;

const toGpsLogsPage = (documents: QueryDocumentSnapshot<DocumentData>[]): GpsLogsPage => ({
  logs: documents.map((document) => ({ ...document.data(), id: document.id }) as GpsLog),
  cursor: documents.length ? documents[documents.length - 1] : null,
  hasMore: documents.length === GPS_LOGS_PAGE_SIZE,
});

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
  onPage: (page: GpsLogsPage) => void,
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

  const logsQuery = query(gpsLogsCollection, orderBy('timestamp', 'desc'), limit(GPS_LOGS_PAGE_SIZE));
  return onSnapshot(
    logsQuery,
    (snapshot) => onPage(toGpsLogsPage(snapshot.docs)),
    (error) => {
      console.error('Failed to subscribe to GPS logs', error);
      onError(error);
    },
  );
};

export const loadMoreGpsLogs = async (cursor: GpsLogCursor): Promise<GpsLogsPage> => {
  await firebaseReady;
  if (!firebaseAuth?.currentUser) {
    throw new Error('يجب تسجيل الدخول إلى Firebase لتحميل سجلات GPS');
  }
  if (!gpsLogsCollection) {
    throw new Error('خدمة Firestore غير مهيأة لتحميل سجلات GPS');
  }

  const olderLogsQuery = query(
    gpsLogsCollection,
    orderBy('timestamp', 'desc'),
    startAfter(cursor),
    limit(GPS_LOGS_PAGE_SIZE),
  );
  const snapshot = await getDocs(olderLogsQuery);
  return toGpsLogsPage(snapshot.docs);
};
