import * as Y from 'yjs';
import { Book, REMOTE, fromBase64, toBase64 } from './book';
import type { Peer } from './presence';
import type { Transport } from './transport';

/**
 * Keeps a Book in sync with everyone else who has it open:
 * - on joining, says what it already has (a state vector); each peer replies with only what the
 *   newcomer is missing plus its own state vector, and the newcomer sends back what that peer is
 *   missing. Only differences travel, so big workbooks don't flood the connection;
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
      const missing = Y.encodeStateAsUpdate(book.doc, fromBase64(msg.sv));
      const sv = toBase64(Y.encodeStateVector(book.doc));
      transport.send({ type: 'state', from: me.id, to: msg.from, data: toBase64(missing), sv });
    } else if (msg.type === 'state' && msg.to === me.id) {
      book.merge(fromBase64(msg.data));
      if (msg.sv) {
        const missing = Y.encodeStateAsUpdate(book.doc, fromBase64(msg.sv));
        transport.send({ type: 'state', from: me.id, to: msg.from, data: toBase64(missing) });
      }
    }
  });
  transport.onPeers(onPeers);

  let current = me;
  transport.setPresence(current);
  void transport.ready.then(() =>
    transport.send({ type: 'hello', from: me.id, sv: toBase64(Y.encodeStateVector(book.doc)) }),
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
