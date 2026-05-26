import { ref, uploadBytesResumable, getDownloadURL } from 'firebase/storage';
import { storage } from './firebase/firebase';

export const uploadLogo = async (file: File, teamName: string): Promise<string> => {
  if (!file) throw new Error("No file provided");
  const ext = file.name.split('.').pop();
  const safeTeamName = teamName.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'unknown_team';
  const fileName = `logos/${safeTeamName}_${Date.now()}.${ext}`;
  
  const storageRef = ref(storage, fileName);
  
  const uploadTask = await uploadBytesResumable(storageRef, file);
  const downloadURL = await getDownloadURL(uploadTask.ref);
  
  return downloadURL;
};
