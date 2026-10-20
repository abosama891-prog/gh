import { addDoc, collection, deleteDoc, doc, getDocs, limit, onSnapshot, orderBy, query, writeBatch } from 'firebase/firestore';
import { firebaseReady, firestore } from './firebase';
import type { FleetRole } from './fleetUserStore';

export type FleetAuditAction = 'edit' | 'delete';

export type FleetAuditEntry = {
  id: string;
  actorName: string;
  actorEmail: string;
  action: FleetAuditAction;
  itemName: string;
  createdAt: string;
};

export type FleetAuditActor = {
  username: string;
  email?: string;
  employeeName?: string;
  guestName?: string;
  role: FleetRole;
  isGuest?: boolean;
};

const auditLogsCollection = firestore ? collection(firestore, 'fleet', 'audit-logs', 'items') : null;

const isExcludedActor = (actor: FleetAuditActor) => {
  const identity = `${actor.username} ${actor.employeeName ?? ''} ${actor.guestName ?? ''} ${actor.email ?? ''}`.toLocaleLowerCase('ar');
  return /\bdeveloper\b|المطور/.test(identity);
};

export const recordFleetAuditEntry = async (
  actor: FleetAuditActor,
  action: FleetAuditAction,
  itemName: string,
): Promise<void> => {
  if (actor.role !== 'admin' || actor.isGuest || isExcludedActor(actor)) return;
  if (!auditLogsCollection) throw new Error('Firestore is not configured');

  await firebaseReady;
  await addDoc(auditLogsCollection, {
    actorName: actor.employeeName || actor.guestName || actor.username || actor.email || 'مسؤول',
    actorEmail: actor.email || '',
    action,
    itemName,
    createdAt: new Date().toISOString(),
  });
};

export const subscribeToFleetAuditLog = (
  callback: (entries: FleetAuditEntry[]) => void,
  onError: (error: Error) => void,
): (() => void) => {
  if (!auditLogsCollection) {
    const error = new Error('Firestore is not configured');
    console.error('Failed to subscribe to fleet audit log', error);
    onError(error);
    return () => {};
  }

  let cancelled = false;
  let unsubscribe = () => {};
  void firebaseReady.then(() => {
    if (cancelled) return;
    unsubscribe = onSnapshot(
      query(auditLogsCollection, orderBy('createdAt', 'desc'), limit(100)),
      (snapshot) => callback(snapshot.docs.map((item) => ({ ...item.data(), id: item.id }) as FleetAuditEntry)),
      (error) => {
        console.error('Failed to listen for fleet audit log', error);
        onError(error);
      },
    );
  }).catch((error: unknown) => {
    const firebaseError = error instanceof Error ? error : new Error(String(error));
    console.error('Failed to initialize fleet audit log', firebaseError);
    onError(firebaseError);
  });

  return () => {
    cancelled = true;
    unsubscribe();
  };
};

export const clearFleetAuditLog = async (): Promise<void> => {
  if (!auditLogsCollection) throw new Error('Firestore is not configured');
  await firebaseReady;
  const snapshot = await getDocs(auditLogsCollection);
  const docs = snapshot.docs;
  // Firestore batch writes are limited to 500 operations per commit,
  // so large audit logs are cleared in chunks.
  const CHUNK_SIZE = 450;
  for (let index = 0; index < docs.length; index += CHUNK_SIZE) {
    const batch = writeBatch(firestore!);
    docs.slice(index, index + CHUNK_SIZE).forEach((document) => batch.delete(document.ref));
    await batch.commit();
  }
};
