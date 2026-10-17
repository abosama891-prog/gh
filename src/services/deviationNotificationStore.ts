import { collection, doc, onSnapshot, orderBy, query, setDoc } from 'firebase/firestore';
import { firebaseReady, firestore } from './firebase';

export type DeviationNotification = {
  kind?: 'deviation';
  id: string;
  vehicleId: string;
  plate: string;
  month: string;
  deviation: number;
  createdAt: string;
};

export type FleetUpdateNotification = {
  kind: 'fleet-update';
  source: 'fuel-entry';
  id: string;
  title: string;
  summary: string;
  createdAt: string;
};

export type FleetNotification = DeviationNotification | FleetUpdateNotification;

const notificationsCollection = firestore ? collection(firestore, 'fleet', 'deviationNotifications', 'items') : null;

export const publishDeviationNotification = async (notification: DeviationNotification): Promise<void> => {
  if (!firestore) throw new Error('Firestore is not configured');
  await firebaseReady;
  await setDoc(doc(firestore, 'fleet', 'deviationNotifications', 'items', notification.id), notification);
};

export const publishFleetUpdateNotification = async (notification: FleetUpdateNotification): Promise<void> => {
  if (!firestore) throw new Error('Firestore is not configured');
  await firebaseReady;
  await setDoc(doc(firestore, 'fleet', 'deviationNotifications', 'items', notification.id), notification);
};

export const subscribeToDeviationNotifications = (
  callback: (notifications: FleetNotification[]) => void,
  onError: (error: Error) => void,
): (() => void) => {
  if (!notificationsCollection) {
    const error = new Error('Firestore is not configured');
    console.error('Failed to subscribe to deviation notifications', error);
    onError(error);
    return () => {};
  }

  let cancelled = false;
  let unsubscribe = () => {};
  void firebaseReady.then(() => {
    if (cancelled) return;
    unsubscribe = onSnapshot(
      query(notificationsCollection, orderBy('createdAt', 'desc')),
      (snapshot) => {
        const notifications = snapshot.docChanges()
          .filter((change) => change.type === 'added')
          .map((change) => {
            const data = change.doc.data();
            return data.kind === 'fleet-update'
              ? data as FleetUpdateNotification
              : { ...data, kind: 'deviation' } as DeviationNotification;
        })
        .filter((notification) => notification.kind !== 'fleet-update' || notification.source === 'fuel-entry');
        callback(notifications);
      },
      (error) => {
        console.error('Failed to listen for deviation notifications', error);
        onError(error);
      },
    );
  }).catch((error: unknown) => {
    const firebaseError = error instanceof Error ? error : new Error(String(error));
    console.error('Failed to initialize deviation notifications', firebaseError);
    onError(firebaseError);
  });

  return () => {
    cancelled = true;
    unsubscribe();
  };
};
