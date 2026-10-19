import { describe, expect, it } from 'vitest';
import type { DocumentPage } from '@ge/contracts';
import { EvidenceWorkspace } from './workspace.js';
import { MemoryEvidenceStore } from './store.js';
import { SourceTraversal, type EvidenceSource } from './source.js';
import { evidenceMount } from './mount.js';

function fixture() {
  const page: DocumentPage = {
    scope: 'document',
    limitations: ['Text only'],
    units: [
      { id: 'section', text: 'Definitions' },
      { id: 'definition', parentId: 'section', text: 'Revenue excludes refunds.' },
      { id: 'target', parentId: 'section', relatedIds: ['definition'], text: 'Revenue was 100.' },
      { id: 'other', text: 'Unrelated appendix' },
    ],
  };
  let reads = 0;
  const source: EvidenceSource = {
    id: 'test-document',
    readPage: async () => {
      reads++;
      return structuredClone(page);
    },
  };
  return { page, source, reads: () => reads };
}

describe('evidence workspace', () => {
  it('publishes only after validated EOF and never marks an early exit complete', async () => {
    const { source } = fixture();
    const traversal = new SourceTraversal(source);
    for await (const _page of traversal.pages()) break;
    expect(() => traversal.receipt).toThrow('EOF');
    const workspace = new EvidenceWorkspace();
    await expect(
      workspace.capture({
        id: 'bad',
        readPage: async () => ({ scope: 'x', units: [], limitations: [], nextCursor: 'same' }),
      }),
    ).rejects.toThrow('cursor');
    expect(await workspace.list()).toEqual([]);
  });
  it('does not publish a cancelled capture', async () => {
    const { source } = fixture();
    const workspace = new EvidenceWorkspace();
    const controller = new AbortController();
    controller.abort();
    await expect(workspace.capture(source, { signal: controller.signal })).rejects.toThrow();
    expect(await workspace.list()).toEqual([]);
  });
  it('rejects duplicate units and enforces source budgets before publishing', async () => {
    const { page, source } = fixture();
    page.units.push(page.units[0]!);
    const workspace = new EvidenceWorkspace();
    await expect(workspace.capture(source)).rejects.toThrow('repeated');
    page.units.pop();
    await expect(workspace.capture(source, { maxBytes: 1 })).rejects.toThrow('budget');
    expect(await workspace.list()).toEqual([]);
  });
  it('reuses immutable data for different tasks and detects changes outside the selection', async () => {
    const { page, source, reads } = fixture();
    const workspace = new EvidenceWorkspace();
    const receipt = await workspace.capture(source);
    const first = await workspace.project(receipt.id, {
      requiredIds: ['target'],
      neighbors: 0,
      relationDepth: 1,
    });
    expect(first.units.map((unit) => unit.id)).toEqual(['target', 'section', 'definition']);
    expect(first.projectionComplete).toBe(false);
    await workspace.project(receipt.id, { query: 'appendix' });
    expect(reads()).toBe(1);
    const unit = await workspace.read(receipt.id, 'target');
    unit.text = 'mutated';
    expect((await workspace.read(receipt.id, 'target')).text).toBe('Revenue was 100.');
    page.units[1]!.text = 'Revenue includes refunds.';
    await expect(workspace.verify(receipt.id, source)).rejects.toThrow('changed');
  });
  it('never truncates required evidence to satisfy a prompt budget', async () => {
    const { page, source } = fixture();
    page.units[2]!.text = '😀"\\'.repeat(1000);
    const workspace = new EvidenceWorkspace();
    const { id } = await workspace.capture(source);
    await expect(
      workspace.project(id, { requiredIds: ['target'], maxBytes: 1024 }),
    ).rejects.toThrow('not truncated');
    const projection = await workspace.project(id, { query: 'Revenue', maxBytes: 1024 });
    expect(new TextEncoder().encode(JSON.stringify(projection)).length).toBeLessThanOrEqual(1024);
  });
  it('binds continuation to the snapshot and query', async () => {
    const { source } = fixture();
    const workspace = new EvidenceWorkspace();
    const { id } = await workspace.capture(source);
    const first = await workspace.query(id, { kind: 'list' }, { limit: 2 });
    expect(first.complete).toBe(false);
    const second = await workspace.query(
      id,
      { kind: 'list' },
      { cursor: first.nextCursor, limit: 2 },
    );
    expect([...first.nodes, ...second.nodes].map((node) => node.id)).toEqual([
      'section',
      'definition',
      'target',
      'other',
    ]);
    expect(second.complete).toBe(true);
    await expect(
      workspace.query(id, { kind: 'grep', literal: 'Revenue' }, { cursor: first.nextCursor }),
    ).rejects.toThrow('different');
  });
  it('searches literal text and expands explicit relationships', async () => {
    const { source } = fixture();
    const workspace = new EvidenceWorkspace();
    const { id } = await workspace.capture(source);
    expect(
      (await workspace.query(id, { kind: 'grep', literal: '100.' })).nodes.map((node) => [
        node.id,
        node.matchOffset,
      ]),
    ).toEqual([['target', 12]]);
    expect((await workspace.query(id, { kind: 'grep', literal: '100.*' })).nodes).toEqual([]);
    expect(
      (await workspace.query(id, { kind: 'expand', id: 'target', direction: 'related' })).nodes.map(
        (node) => node.id,
      ),
    ).toEqual(['definition']);
    expect(
      (await workspace.query(id, { kind: 'expand', id: 'section', direction: 'children' })).nodes,
    ).toHaveLength(2);
  });
  it('reads oversized Unicode nodes incrementally without losing text', async () => {
    const { source, page } = fixture();
    page.units[2]!.text = 'A😀B😀C';
    const workspace = new EvidenceWorkspace();
    const { id } = await workspace.capture(source);
    let offset: number | undefined = 0;
    let output = '';
    do {
      const chunk = await workspace.readText(id, 'target', { offset, maxBytes: 4 });
      output += chunk.text;
      offset = chunk.nextOffset;
    } while (offset !== undefined);
    expect(output).toBe(page.units[2]!.text);
    await expect(workspace.readText(id, 'target', { offset: 2 })).rejects.toThrow('Unicode');
  });
  it('lets existing file-read commands follow oversized text to EOF', async () => {
    const { source, page } = fixture();
    page.units[2]!.text = '😀'.repeat(20000);
    const workspace = new EvidenceWorkspace();
    const { id } = await workspace.capture(source);
    const mount = evidenceMount(workspace);
    let path: string | undefined = `${id}/windows/2/0.json`;
    let text = '';
    while (path) {
      const view = await mount.readFile(path);
      expect(view?.truncated).toBe(false);
      const window = JSON.parse(view!.text);
      text += window.text;
      path = window.nextPath?.slice('/evidence/'.length);
    }
    expect(text).toBe(page.units[2]!.text);
  });
  it('rejects a tampered persistent store and out-of-snapshot retrieval', async () => {
    const { source } = fixture();
    const store = new MemoryEvidenceStore();
    const workspace = new EvidenceWorkspace(store);
    const { id } = await workspace.capture(source);
    const snapshot = await workspace.load(id);
    snapshot.pages[0]!.units[0]!.text = 'tampered';
    const corrupt = new EvidenceWorkspace({
      ...store,
      get: async () => snapshot,
      put: async () => {},
      delete: async () => {},
      list: async () => [id],
    });
    await expect(corrupt.load(id)).rejects.toThrow('integrity');
    const badSelector = new EvidenceWorkspace(store, async () => ['not-present']);
    await expect(badSelector.project(id)).rejects.toThrow('outside');
  });
  it('enforces explicit lifecycle quotas and path-scoped literal DocFs search', async () => {
    const { source } = fixture();
    const workspace = new EvidenceWorkspace(new MemoryEvidenceStore(100000, 1));
    const { id } = await workspace.capture(source);
    await expect(workspace.capture({ ...source, id: 'another' })).rejects.toThrow('quota');
    const mount = evidenceMount(workspace);
    expect(await mount.search!(`${id}/units/0/3.json`, 'Revenue', {})).toEqual([]);
    expect(await mount.search!(`${id}/units/0/2.json`, 'Revenue', {})).toHaveLength(1);
    const tiny = await mount.readFile(`${id}/units/0/2.json`, { maxBytes: 4 });
    const stat = await mount.stat(`${id}/units/0/2.json`);
    expect(tiny?.truncated).toBe(true);
    expect(stat!.size).toBeGreaterThan(tiny!.bytes);
    await workspace.release(id);
    expect(await workspace.list()).toEqual([]);
  });
});
