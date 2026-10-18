import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import { fromBase64, toBase64 } from '../collab/book';
import type { Store } from './store';

const url = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

export const supabase: SupabaseClient | null = url && anonKey ? createClient(url, anonKey) : null;
export const supabaseUrl = url ?? '';

/**
 * Saves in Supabase (tables in supabase/schema.sql). Needs a signed-in user, and only
 * returns the books they own or were invited to.
 */
export function cloudStore(client: SupabaseClient): Store {
  return {
    kind: 'cloud',

    async list() {
      const { data, error } = await client
        .from('workbooks')
        .select('id, name, updated_at')
        .order('updated_at', { ascending: false });
      if (error) throw error;
      return (data ?? []).map((r) => ({ id: r.id, name: r.name, updatedAt: r.updated_at }));
    },

    async load(id) {
      const { data, error } = await client.from('workbooks').select('data, ydoc, owner').eq('id', id).maybeSingle();
      if (error) throw error;
      return data ? { wb: data.data, state: data.ydoc ?? undefined, owner: data.owner } : null;
    },

    async save(book) {
      const id = book.snapshot.id;
      // Read-merge-write: whatever a collaborator saved since we last synced is kept.
      const { data: current, error: readError } = await client.from('workbooks').select('ydoc').eq('id', id).maybeSingle();
      if (readError) throw readError;
      if (current?.ydoc) book.merge(fromBase64(current.ydoc));
      const wb = book.snapshot;
      const row = { name: wb.name, data: wb, ydoc: toBase64(book.encode()), updated_at: wb.updatedAt || new Date().toISOString() };
      const { error } = current
        ? await client.from('workbooks').update(row).eq('id', id)
        : await client.from('workbooks').insert({ id, ...row });
      if (error) throw error;
    },

    async remove(id) {
      const { error } = await client.from('workbooks').delete().eq('id', id);
      if (error) throw error;
    },
  };
}

export async function listMembers(client: SupabaseClient, workbookId: string): Promise<string[]> {
  const { data, error } = await client.from('workbook_members').select('email').eq('workbook_id', workbookId).order('email');
  if (error) throw error;
  return (data ?? []).map((r) => r.email);
}

export async function addMember(client: SupabaseClient, workbookId: string, email: string) {
  const { error } = await client.from('workbook_members').insert({ workbook_id: workbookId, email: email.trim().toLowerCase() });
  if (error && error.code !== '23505') throw error; // already invited
}

export async function removeMember(client: SupabaseClient, workbookId: string, email: string) {
  const { error } = await client.from('workbook_members').delete().eq('workbook_id', workbookId).eq('email', email);
  if (error) throw error;
}
