import { collection, addDoc, getDocs, doc, updateDoc, query, orderBy, getDocFromServer, deleteDoc } from 'firebase/firestore';
import { db } from './firebase/firebase';
import { Player, ApplicationDraft } from './sheets';

export interface FirestoreSubmission {
  id?: string;
  teamName: string;
  zone: string;
  captainName: string;
  captainPhone: string;
  logoUrl: string | null;
  players: Player[];
  createdAt: string;
  status: 'pending' | 'approved' | 'rejected';
  synced: boolean;
  stage?: 'qualifier' | 'final';
}

const SUBMISSIONS_COLLECTION = 'submissions';

/**
 * Validates the Firestore connection on boot
 */
export async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.warn("Please check your Firebase configuration: Firestore appears offline.");
    }
  }
}

/**
 * Saves a new application submission to Firestore
 */
export async function saveSubmissionToFirestore(draft: ApplicationDraft): Promise<string> {
  try {
    const submissionData: Omit<FirestoreSubmission, 'id'> = {
      teamName: draft.teamName,
      zone: draft.zone,
      captainName: draft.captainName,
      captainPhone: draft.captainPhone,
      logoUrl: draft.logoUrl,
      players: draft.players,
      createdAt: new Date().toISOString(),
      status: 'pending',
      synced: false,
      stage: draft.stage || 'qualifier'
    };

    const docRef = await addDoc(collection(db, SUBMISSIONS_COLLECTION), submissionData);
    return docRef.id;
  } catch (err) {
    console.error("Failed to save submission to Firestore:", err);
    throw new Error(`Ошибка сохранения в Firestore: ${err instanceof Error ? err.message : String(err)}`);
  }
}

/**
 * Fetches the latest submission for a specific team and stage
 */
export async function getTeamSubmission(teamName: string, stage: 'qualifier' | 'final', zone?: string): Promise<FirestoreSubmission | null> {
  try {
    const submissionsRef = collection(db, SUBMISSIONS_COLLECTION);
    const q = query(submissionsRef, orderBy('createdAt', 'desc'));
    const querySnapshot = await getDocs(q);
    
    // We filter client-side to avoid needing a composite index
    for (const doc of querySnapshot.docs) {
      const data = doc.data() as FirestoreSubmission;
      if (!data.teamName) continue;
      if (data.teamName.trim().toLowerCase() === teamName.trim().toLowerCase() && (data.stage || 'qualifier') === stage) {
         return { id: doc.id, ...data };
      }
    }
    return null;
  } catch (err) {
    console.error("Failed to fetch team submission:", err);
    return null;
  }
}

/**
 * Fetches all submissions from Firestore (ordered by creation date desc)
 */
export async function getSubmissionsFromFirestore(): Promise<FirestoreSubmission[]> {
  try {
    const q = query(collection(db, SUBMISSIONS_COLLECTION), orderBy('createdAt', 'desc'));
    const snapshot = await getDocs(q);
    return snapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data()
    } as FirestoreSubmission));
  } catch (err) {
    console.error("Failed to get submissions from Firestore:", err);
    throw new Error(`Ошибка получения из Firestore: ${err instanceof Error ? err.message : String(err)}`);
  }
}

/**
 * Updates the sync/status flag of a submission
 */
export async function updateSubmissionData(id: string, updates: Partial<FirestoreSubmission>): Promise<void> {
  try {
    const docRef = doc(db, SUBMISSIONS_COLLECTION, id);
    await updateDoc(docRef, updates);
  } catch (err) {
    console.error("Failed to update submission data in Firestore:", err);
    throw new Error(`Ошибка обновления Firestore: ${err instanceof Error ? err.message : String(err)}`);
  }
}

/**
 * Deletes a submission from Firestore
 */
export async function deleteSubmissionFromFirestore(id: string): Promise<void> {
  try {
    const docRef = doc(db, SUBMISSIONS_COLLECTION, id);
    await deleteDoc(docRef);
  } catch (err) {
    console.error("Failed to delete submission from Firestore:", err);
    throw new Error(`Ошибка удаления из Firestore: ${err instanceof Error ? err.message : String(err)}`);
  }
}
