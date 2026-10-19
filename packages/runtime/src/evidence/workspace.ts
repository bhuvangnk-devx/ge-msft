import {
  ContextProjectionRequestSchema,
  EvidenceSnapshotSchema,
  type ContextProjection,
  type ContextProjectionRequest,
  type DocumentPage,
  type EvidenceSnapshot,
  type SourceLimits,
  type SourceReceipt,
} from '@ge/contracts';
import { evidenceHash } from './hash.js';
import { navigateEvidence, type EvidenceQuery } from './navigation.js';
import { byteLength } from '../docfs/bytes.js';
import { SourceTraversal, verifySource, type EvidenceSource } from './source.js';
import { MemoryEvidenceStore, type EvidenceStore } from './store.js';

export function evidencePath(snapshot: string, index: number): string {
  return `/evidence/${snapshot}/units/${Math.floor(index / 100)}/${index}.json`;
}
export function snapshotReceipt(snapshot: EvidenceSnapshot): SourceReceipt {
  const { pages: _pages, ...receipt } = snapshot;
  return receipt;
}

/** Retrieval policy is replaceable; returned IDs must belong to the captured source. */
export type EvidenceSelector = (
  query: string,
  units: ReadonlyArray<Readonly<DocumentPage['units'][number]>>,
) => Promise<readonly string[]>;
export const lexicalEvidenceSelector: EvidenceSelector = async (query, units) => {
  const terms = [...new Set(query.toLocaleLowerCase().match(/[\p{L}\p{N}_]+/gu) ?? [])];
  return units
    .map((unit, index) => {
      const text = unit.text.toLocaleLowerCase();
      return { id: unit.id, index, score: terms.filter((term) => text.includes(term)).length };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map((item) => item.id);
};

/** Source-neutral evidence, independent of Office, model prompts, task verbs and write authority. */
export class EvidenceWorkspace {
  constructor(
    private readonly store: EvidenceStore = new MemoryEvidenceStore(),
    private readonly selector: EvidenceSelector = lexicalEvidenceSelector,
  ) {}
  async capture(
    source: EvidenceSource,
    options: Partial<SourceLimits> & { signal?: AbortSignal } = {},
  ): Promise<SourceReceipt> {
    const { signal, ...limits } = options;
    const traversal = new SourceTraversal(source, limits);
    const pages: DocumentPage[] = [];
    for await (const page of traversal.pages(signal)) pages.push(page);
    signal?.throwIfAborted();
    const receipt = traversal.receipt;
    // Nothing is published until validated EOF. A failed/cancelled traversal leaves no partial snapshot.
    await this.store.put({ ...receipt, pages });
    return receipt;
  }
  async list(): Promise<string[]> {
    return this.store.list();
  }
  async release(id: string): Promise<void> {
    await this.store.delete(id);
  }
  async load(id: string): Promise<EvidenceSnapshot> {
    const raw = await this.store.get(id);
    if (raw === undefined) throw new Error('Evidence snapshot unavailable or released.');
    const snapshot = EvidenceSnapshotSchema.parse(raw);
    if (snapshot.id !== id) throw new Error('Evidence store returned the wrong snapshot.');
    // Treat storage as an untrusted boundary, including a caller-supplied persistent adapter.
    let index = 0;
    const traversal = new SourceTraversal(
      {
        id: snapshot.sourceId,
        readPage: async () => {
          const page = snapshot.pages[index++];
          if (!page) throw new Error('Stored evidence ended before EOF.');
          return page;
        },
      },
      { maxBytes: 256 * 1024 * 1024 },
    );
    for await (const _page of traversal.pages()) void _page;
    const { id: _id, pages: _pages, ...body } = snapshot;
    if (
      index !== snapshot.pages.length ||
      traversal.receipt.id !== snapshot.id ||
      (await evidenceHash(JSON.stringify(body))) !== snapshot.id
    )
      throw new Error('Evidence integrity check failed.');
    return snapshot;
  }
  async verify(id: string, source: EvidenceSource, signal?: AbortSignal): Promise<void> {
    await verifySource(source, snapshotReceipt(await this.load(id)), signal);
  }
  async read(id: string, unitId: string): Promise<DocumentPage['units'][number]> {
    const unit = (await this.load(id)).pages
      .flatMap((page) => page.units)
      .find((unit) => unit.id === unitId);
    if (!unit) throw new Error('Evidence unit does not exist in this snapshot.');
    return structuredClone(unit);
  }
  async query(
    id: string,
    query: EvidenceQuery,
    options: { cursor?: string; limit?: number; maxBytes?: number } = {},
  ) {
    return navigateEvidence(await this.load(id), query, options);
  }
  /** Exact text windows; callers can advance nextOffset without losing oversized nodes. */
  async readText(id: string, unitId: string, options: { offset?: number; maxBytes?: number } = {}) {
    const unit = await this.read(id, unitId);
    const offset = options.offset ?? 0;
    const maxBytes = options.maxBytes ?? 16 * 1024;
    if (
      !Number.isSafeInteger(offset) ||
      offset < 0 ||
      offset > unit.text.length ||
      !Number.isSafeInteger(maxBytes) ||
      maxBytes < 4 ||
      maxBytes > 96 * 1024
    )
      throw new Error('Invalid evidence text window.');
    if (
      offset > 0 &&
      /[\uDC00-\uDFFF]/.test(unit.text[offset] ?? '') &&
      /[\uD800-\uDBFF]/.test(unit.text[offset - 1]!)
    )
      throw new Error('Text offset splits a Unicode character.');
    let text = '';
    let bytes = 0;
    for (const character of unit.text.slice(offset)) {
      const size = byteLength(character);
      if (bytes + size > maxBytes) break;
      text += character;
      bytes += size;
    }
    const next = offset + text.length;
    return {
      snapshotId: id,
      id: unitId,
      offset,
      text,
      nextOffset: next < unit.text.length ? next : undefined,
      complete: next === unit.text.length,
    };
  }
  async project(id: string, input: ContextProjectionRequest = {}): Promise<ContextProjection> {
    const request = ContextProjectionRequestSchema.parse(input);
    const snapshot = await this.load(id);
    const units = snapshot.pages.flatMap((page) => page.units);
    const indexes = new Map(units.map((unit, index) => [unit.id, index]));
    const required = [...new Set(request.requiredIds)];
    for (const id of required)
      if (!indexes.has(id)) throw new Error(`Required evidence is missing: ${id}`);
    const candidates = new Set(required);
    const unresolved = new Set<string>();
    let frontier = required;
    for (let depth = 0; depth < request.relationDepth; depth++) {
      const next: string[] = [];
      for (const id of frontier) {
        const unit = units[indexes.get(id)!]!;
        for (const related of [
          ...(unit.parentId ? [unit.parentId] : []),
          ...(unit.relatedIds ?? []),
        ]) {
          if (!indexes.has(related)) {
            unresolved.add(related);
            continue;
          }
          if (!candidates.has(related)) {
            candidates.add(related);
            next.push(related);
          }
        }
      }
      frontier = next;
    }
    for (const id of required) {
      const index = indexes.get(id)!;
      for (
        let n = Math.max(0, index - request.neighbors);
        n <= Math.min(units.length - 1, index + request.neighbors);
        n++
      ) {
        if (units[n]!.parentId === units[index]!.parentId) candidates.add(units[n]!.id);
      }
    }
    const selected = await this.selector(request.query, structuredClone(units));
    for (const id of selected) {
      if (!indexes.has(id)) throw new Error('Selector returned evidence outside the snapshot.');
      candidates.add(id);
    }
    const projection: ContextProjection = {
      snapshotId: snapshot.id,
      sourceId: snapshot.sourceId,
      scope: snapshot.scope,
      sourceComplete: true,
      projectionComplete: false,
      selection:
        this.selector === lexicalEvidenceSelector
          ? 'required-structure-lexical'
          : 'required-structure-custom',
      omittedUnits: units.length,
      unresolvedRelations: unresolved.size,
      limitations: [...snapshot.limitations],
      units: [],
    };
    if (byteLength(JSON.stringify(projection)) > request.maxBytes)
      throw new Error('Projection metadata exceeds its byte budget.');
    const requiredSet = new Set(required);
    let encodedBytes = byteLength(JSON.stringify(projection));
    for (const id of candidates) {
      const index = indexes.get(id)!;
      const unit = units[index]!;
      const entry = {
        id,
        hash: '0'.repeat(64),
        path: evidencePath(snapshot.id, index),
        text: unit.text,
      };
      const nextOmitted = projection.omittedUnits - 1;
      const delta =
        byteLength(JSON.stringify(entry)) +
        (projection.units.length ? 1 : 0) +
        String(nextOmitted).length -
        String(projection.omittedUnits).length +
        (nextOmitted === 0 ? -1 : 0); // JSON true is one byte shorter than false.
      if (encodedBytes + delta > request.maxBytes) {
        if (requiredSet.has(id))
          throw new Error('Required evidence exceeds the projection budget; it was not truncated.');
        continue;
      }
      entry.hash = await evidenceHash(JSON.stringify(unit));
      projection.units.push(entry);
      projection.omittedUnits = nextOmitted;
      projection.projectionComplete = nextOmitted === 0;
      encodedBytes += delta;
    }
    return projection;
  }
}
