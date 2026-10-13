import { collection, doc, limit, onSnapshot, orderBy, query, setDoc } from 'firebase/firestore';
import { firebaseReady, firestore } from './firebase';

export type DeviationNotification = {
  id: string;
  vehicleId: string;
  plate: string;
  month: string;
  deviation: number;
  createdAt: string;
};

const notificationsCollection = firestore ? collection(firestore, 'fleet', 'deviationNotifications', 'items') : null;

export const publishDeviationNotification = async (notification: DeviationNotification): Promise<void> => {
  if (!firestore) throw new Error('Firestore is not configured');
  await firebaseReady;
  await setDoc(doc(firestore, 'fleet', 'deviationNotifications', 'items', notification.id), notification);
};

export const subscribeToDeviationNotifications = (
  callback: (notifications: DeviationNotification[]) => void,
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
    let initialSnapshot = true;
    unsubscribe = onSnapshot(
      query(notificationsCollection, orderBy('createdAt', 'desc'), limit(100)),
      (snapshot) => {
        if (initialSnapshot) {
          initialSnapshot = false;
          return;
        }
        const notifications = snapshot.docChanges()
          .filter((change) => change.type === 'added')
          .map((change) => change.doc.data() as DeviationNotification);
        if (notifications.length) callback(notifications);
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
