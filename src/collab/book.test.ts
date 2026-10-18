import { describe, expect, it } from 'vitest';
import { newWorkbook, type Workbook } from '../model/types';
import { Book } from './book';
import type { Peer } from './presence';
import { connect } from './provider';
import type { Message, Transport } from './transport';

function base(): Workbook {
  const wb = newWorkbook('Prueba');
  wb.sheets[0].cells = { A1: { raw: '1' }, B1: { raw: '=A1*2' } };
  return wb;
}

/** Copies a's state into b and b's into a, like two people syncing. */
function sync(a: Book, b: Book) {
  a.merge(b.encode());
  b.merge(a.encode());
}

const cells = (b: Book) => b.snapshot.sheets[0].cells;

describe('Book', () => {
  it('two people converting the same saved workbook end with one copy, not two', () => {
    const wb = base();
    const a = Book.fromWorkbook(wb);
    const b = Book.fromWorkbook(wb);
    sync(a, b);
    expect(a.snapshot.sheets).toHaveLength(1);
    expect(b.snapshot).toEqual(a.snapshot);
  });

  it('merges simultaneous edits to different cells', () => {
    const wb = base();
    const a = Book.fromWorkbook(wb);
    const b = Book.fromWorkbook(wb);
    const sheetId = wb.sheets[0].id;
    a.edit({ type: 'setCell', sheetId, key: 'A2', raw: 'de Ana' });
    b.edit({ type: 'setCell', sheetId, key: 'C3', raw: 'de Beto' });
    b.edit({ type: 'addSheet' });
    sync(a, b);
    expect(cells(a).A2.raw).toBe('de Ana');
    expect(cells(a).C3.raw).toBe('de Beto');
    expect(a.snapshot.sheets).toHaveLength(2);
    expect(b.snapshot).toEqual(a.snapshot);
  });

  it('settles simultaneous edits to the same cell the same way for everyone', () => {
    const wb = base();
    const a = Book.fromWorkbook(wb);
    const b = Book.fromWorkbook(wb);
    const sheetId = wb.sheets[0].id;
    a.edit({ type: 'setCell', sheetId, key: 'A1', raw: '10' });
    b.edit({ type: 'setCell', sheetId, key: 'A1', raw: '20' });
    sync(a, b);
    expect(['10', '20']).toContain(cells(a).A1.raw);
    expect(cells(b).A1.raw).toBe(cells(a).A1.raw);
  });

  it('undo only reverts my own changes', () => {
    const wb = base();
    const a = Book.fromWorkbook(wb);
    const b = Book.fromWorkbook(wb);
    const sheetId = wb.sheets[0].id;
    a.edit({ type: 'setCell', sheetId, key: 'A2', raw: 'mío' });
    b.edit({ type: 'setCell', sheetId, key: 'A3', raw: 'ajeno' });
    sync(a, b);
    a.undo();
    expect(cells(a).A2).toBeUndefined();
    expect(cells(a).A3.raw).toBe('ajeno');
    a.redo();
    expect(cells(a).A2.raw).toBe('mío');
  });
});

/** All transports connected to one in-memory "room", delivering synchronously. */
function room() {
  const members: { id: string; deliver(m: Message): void }[] = [];
  return (id: string): Transport => {
    let onMsg: (m: Message) => void = () => {};
    members.push({ id, deliver: (m) => onMsg(m) });
    return {
      ready: Promise.resolve(),
      send: (m) => members.filter((x) => x.id !== id).forEach((x) => x.deliver(m)),
      onMessage: (fn) => (onMsg = fn),
      setPresence: () => {},
      onPeers: () => {},
      close: () => {},
    };
  };
}

const peer = (id: string): Peer => ({ id, name: id, color: '#fff' });

describe('connect', () => {
  it('only sends differences when someone joins', async () => {
    const sent: Message[] = [];
    const join = room();
    const wb = base();
    const a = Book.fromWorkbook(wb);
    connect(a, join('ana'), peer('ana'), () => {});
    await Promise.resolve();
    const b = Book.fromWorkbook(wb); // same saved copy: nothing to exchange
    const t = join('beto');
    const send = t.send;
    t.send = (m) => {
      sent.push(m);
      send(m);
    };
    connect(b, t, peer('beto'), () => {});
    await Promise.resolve();
    const bytes = sent.reduce((n, m) => n + JSON.stringify(m).length, 0);
    expect(bytes).toBeLessThan(400);
  });

  it('brings a late joiner up to date and then streams every change', async () => {
    const join = room();
    const wb = base();
    const sheetId = wb.sheets[0].id;
    const a = Book.fromWorkbook(wb);
    a.edit({ type: 'setCell', sheetId, key: 'A5', raw: 'antes de que llegue Beto' });
    connect(a, join('ana'), peer('ana'), () => {});
    await Promise.resolve();

    // Beto opens an older saved copy, without Ana's last edit.
    const b = Book.fromWorkbook(wb);
    b.edit({ type: 'setCell', sheetId, key: 'D1', raw: 'offline de Beto' });
    connect(b, join('beto'), peer('beto'), () => {});
    await Promise.resolve();
    await Promise.resolve();

    expect(cells(b).A5.raw).toBe('antes de que llegue Beto');
    expect(cells(a).D1.raw).toBe('offline de Beto');

    a.edit({ type: 'setCell', sheetId, key: 'B2', raw: 'en vivo' });
    expect(cells(b).B2.raw).toBe('en vivo');
    b.undo(); // Beto's own undo removes his offline edit, everywhere
    expect(cells(a).D1).toBeUndefined();
    expect(cells(a).B2.raw).toBe('en vivo');
  });
});
