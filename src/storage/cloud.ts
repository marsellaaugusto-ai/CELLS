import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import type { Workbook } from '../model/types';
import type { Store } from './store';

const url = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

export const supabase: SupabaseClient | null = url && anonKey ? createClient(url, anonKey) : null;
export const supabaseUrl = url ?? '';

interface Row {
  id: string;
  name: string;
  data: Workbook;
  updated_at: string;
}

/** Saves in Supabase (table "workbooks", see supabase/schema.sql). Needs a signed-in user. */
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
      const { data, error } = await client.from('workbooks').select('data').eq('id', id).maybeSingle();
      if (error) throw error;
      return (data?.data as Workbook | undefined) ?? null;
    },

    async save(wb) {
      const { error } = await client
        .from('workbooks')
        .upsert({ id: wb.id, name: wb.name, data: wb, updated_at: wb.updatedAt });
      if (error) throw error;
    },

    async remove(id) {
      const { error } = await client.from('workbooks').delete().eq('id', id);
      if (error) throw error;
    },

    subscribe(id, onChange) {
      const channel = client
        .channel(`workbook:${id}`)
        .on(
          'postgres_changes',
          { event: 'UPDATE', schema: 'public', table: 'workbooks', filter: `id=eq.${id}` },
          (payload) => onChange((payload.new as Row).data),
        )
        .subscribe();
      return () => {
        void client.removeChannel(channel);
      };
    },
  };
}
