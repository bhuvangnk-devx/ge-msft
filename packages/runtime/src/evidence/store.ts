import type { EvidenceSnapshot } from '@ge/contracts';
import { byteLength } from '../docfs/bytes.js';

/** Atomic publication port. Implementations must isolate principals and enforce aggregate quotas.
 * Persist only source data; never credentials, approval grants, or executable callbacks. */
export interface EvidenceStore {
  get(id: string): Promise<unknown | undefined>;
  put(snapshot: EvidenceSnapshot): Promise<void>;
  delete(id: string): Promise<void>;
  list(): Promise<string[]>;
}

/** No implicit eviction: an outstanding reference must not silently point at different evidence. */
export class MemoryEvidenceStore implements EvidenceStore {
  private readonly entries = new Map<string, string>();
  private bytes = 0;
  constructor(
    private readonly maxBytes = 128 * 1024 * 1024,
    private readonly maxSnapshots = 8,
  ) {
    if (
      !Number.isSafeInteger(maxBytes) ||
      maxBytes < 1 ||
      !Number.isSafeInteger(maxSnapshots) ||
      maxSnapshots < 1
    )
      throw new Error('Evidence store quotas must be positive safe integers.');
  }
  async get(id: string): Promise<unknown | undefined> {
    const entry = this.entries.get(id);
    return entry === undefined ? undefined : JSON.parse(entry);
  }
  async put(snapshot: EvidenceSnapshot): Promise<void> {
    const encoded = JSON.stringify(snapshot);
    const prior = this.entries.get(snapshot.id);
    if (prior !== undefined && prior !== encoded) throw new Error('Immutable snapshot conflict.');
    if (prior !== undefined) return;
    const bytes = byteLength(encoded);
    if (this.entries.size >= this.maxSnapshots || this.bytes + bytes > this.maxBytes)
      throw new Error('Evidence store quota exceeded. Release an unused snapshot explicitly.');
    this.entries.set(snapshot.id, encoded);
    this.bytes += bytes;
  }
  async delete(id: string): Promise<void> {
    const prior = this.entries.get(id);
    if (prior !== undefined) {
      this.bytes -= byteLength(prior);
      this.entries.delete(id);
    }
  }
  async list(): Promise<string[]> {
    return [...this.entries.keys()];
  }
}
