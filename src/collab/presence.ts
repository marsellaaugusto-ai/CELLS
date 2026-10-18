import { newId } from '../model/types';

export interface Peer {
  id: string;
  name: string;
  color: string;
  /** Where this person's cursor is. */
  at?: { sheetId: string; r: number; c: number };
}

// Bright enough to read as a ring on the dark tiles, and distinct from the selection blue.
const COLORS = ['#e8743b', '#5cc26a', '#e9b326', '#a77de0', '#2bb5c9', '#ff6b9a', '#9ccf3e', '#f0a35e'];

function guestName() {
  const KEY = 'cells:guest-name';
  try {
    const saved = localStorage.getItem(KEY);
    if (saved) return saved;
    const name = `Invitado ${Math.floor(100 + Math.random() * 900)}`;
    localStorage.setItem(KEY, name);
    return name;
  } catch {
    return 'Invitado';
  }
}

/** Every open tab is its own participant, so two tabs show up as two people. */
export function makeMe(email: string | null): Peer {
  const id = newId();
  let h = 0;
  for (const ch of id) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return { id, name: email ? email.split('@')[0] : guestName(), color: COLORS[h % COLORS.length] };
}
