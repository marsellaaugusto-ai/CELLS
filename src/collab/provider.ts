import * as Y from 'yjs';
import { Book, REMOTE, fromBase64, toBase64 } from './book';
import type { Peer } from './presence';
import type { Transport } from './transport';

/**
 * Keeps a Book in sync with everyone else who has it open:
 * - on joining, sends its whole state plus what it already knows, and each peer replies
 *   with only what the newcomer is missing;
 * - after that, every local change is sent as a small Yjs update.
 */
export function connect(book: Book, transport: Transport, me: Peer, onPeers: (peers: Peer[]) => void) {
  const onUpdate = (update: Uint8Array, origin: unknown) => {
    if (origin !== REMOTE) transport.send({ type: 'update', from: me.id, data: toBase64(update) });
  };
  book.doc.on('update', onUpdate);

  transport.onMessage((msg) => {
    if (msg.from === me.id) return;
    if (msg.type === 'update') book.merge(fromBase64(msg.data));
    else if (msg.type === 'hello') {
      book.merge(fromBase64(msg.data));
      const missing = Y.encodeStateAsUpdate(book.doc, fromBase64(msg.sv));
      transport.send({ type: 'state', from: me.id, to: msg.from, data: toBase64(missing) });
    } else if (msg.type === 'state' && msg.to === me.id) book.merge(fromBase64(msg.data));
  });
  transport.onPeers(onPeers);

  let current = me;
  transport.setPresence(current);
  void transport.ready.then(() =>
    transport.send({
      type: 'hello',
      from: me.id,
      sv: toBase64(Y.encodeStateVector(book.doc)),
      data: toBase64(book.encode()),
    }),
  ).catch(() => onPeers([]));

  return {
    moveCursor(at: Peer['at']) {
      current = { ...current, at };
      transport.setPresence(current);
    },
    close() {
      book.doc.off('update', onUpdate);
      transport.close();
    },
  };
}
