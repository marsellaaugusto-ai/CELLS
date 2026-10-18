import type { SupabaseClient } from '@supabase/supabase-js';
import type { Peer } from './presence';

/** Messages exchanged between copies of the same workbook. Binary Yjs data travels as base64. */
export type Message =
  | { type: 'update'; from: string; data: string }
  | { type: 'hello'; from: string; sv: string }
  /** What the receiver is missing; `sv` asks the receiver to send back what the sender is missing. */
  | { type: 'state'; from: string; to: string; data: string; sv?: string };

export interface Transport {
  send(msg: Message): void;
  onMessage(fn: (msg: Message) => void): void;
  /** Announces who I am and where my cursor is. */
  setPresence(me: Peer): void;
  onPeers(fn: (peers: Peer[]) => void): void;
  /** Resolves once the connection can carry messages. */
  ready: Promise<void>;
  close(): void;
}

type PresenceMsg = { type: 'presence'; peer: Peer } | { type: 'bye'; id: string };

/**
 * Tabs of the same browser, without any server. Lets you try collaboration locally
 * and keeps two open tabs of the same workbook in sync.
 */
export function tabTransport(workbookId: string): Transport {
  const bc = new BroadcastChannel(`cells:${workbookId}`);
  const seen = new Map<string, { peer: Peer; at: number }>();
  let onMsg: (m: Message) => void = () => {};
  let onPeers: (p: Peer[]) => void = () => {};
  let me: Peer | null = null;

  const emitPeers = () => onPeers([...seen.values()].map((s) => s.peer));
  bc.onmessage = (e: MessageEvent<Message | PresenceMsg>) => {
    const m = e.data;
    if (m.type === 'presence') {
      const isNew = !seen.has(m.peer.id);
      seen.set(m.peer.id, { peer: m.peer, at: Date.now() });
      emitPeers();
      if (isNew && me) bc.postMessage({ type: 'presence', peer: me } satisfies PresenceMsg);
    } else if (m.type === 'bye') {
      seen.delete(m.id);
      emitPeers();
    } else onMsg(m);
  };

  const heartbeat = setInterval(() => {
    if (me) bc.postMessage({ type: 'presence', peer: me } satisfies PresenceMsg);
    const cutoff = Date.now() - 10_000;
    let changed = false;
    for (const [id, s] of seen) if (s.at < cutoff) changed = seen.delete(id) || changed;
    if (changed) emitPeers();
  }, 4_000);

  const bye = () => me && bc.postMessage({ type: 'bye', id: me.id } satisfies PresenceMsg);
  window.addEventListener('pagehide', bye);

  return {
    ready: Promise.resolve(),
    send: (msg) => bc.postMessage(msg),
    onMessage: (fn) => (onMsg = fn),
    setPresence(peer) {
      me = peer;
      bc.postMessage({ type: 'presence', peer } satisfies PresenceMsg);
    },
    onPeers: (fn) => (onPeers = fn),
    close() {
      bye();
      clearInterval(heartbeat);
      window.removeEventListener('pagehide', bye);
      bc.close();
    },
  };
}

/**
 * Supabase Realtime: people on different devices. The channel is private, so only
 * users with access to the workbook can join (policies in supabase/schema.sql).
 */
export function cloudTransport(client: SupabaseClient, workbookId: string, meId: string): Transport {
  const channel = client.channel(`wb:${workbookId}`, {
    config: { private: true, broadcast: { self: false }, presence: { key: meId } },
  });
  let onMsg: (m: Message) => void = () => {};
  let onPeers: (p: Peer[]) => void = () => {};
  let latest: Peer | null = null;
  let subscribed = false;

  channel.on('broadcast', { event: 'y' }, ({ payload }) => onMsg(payload as Message));
  channel.on('presence', { event: 'sync' }, () => {
    const state = channel.presenceState<Peer>();
    onPeers(Object.values(state).map((list) => list[list.length - 1]).filter((p) => p.id !== meId));
  });

  const ready = (async () => {
    await client.realtime.setAuth();
    await new Promise<void>((resolve, reject) =>
      channel.subscribe((status, err) => {
        if (status === 'SUBSCRIBED') {
          subscribed = true;
          if (latest) void channel.track(latest);
          resolve();
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') reject(err ?? new Error(status));
      }),
    );
  })();

  return {
    ready,
    send: (msg) => void channel.send({ type: 'broadcast', event: 'y', payload: msg }),
    onMessage: (fn) => (onMsg = fn),
    setPresence(peer) {
      latest = peer;
      if (subscribed) void channel.track(peer);
    },
    onPeers: (fn) => (onPeers = fn),
    close() {
      void client.removeChannel(channel);
    },
  };
}
