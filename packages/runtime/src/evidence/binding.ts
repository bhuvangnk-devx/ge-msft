import type { ActuationRequest, ContextProjection } from '@ge/contracts';
import type { EvidenceSource } from './source.js';
import type { EvidenceWorkspace } from './workspace.js';

/** A trusted, process-local policy binding. It is never serialized with a source snapshot. */
export class EvidenceBinding {
  private invalidated = false;
  private readonly projection: ContextProjection;
  constructor(
    projection: ContextProjection,
    private readonly workspace: EvidenceWorkspace,
    private readonly source: EvidenceSource,
    private readonly authorizeEffect: (request: ActuationRequest) => boolean = () => false,
  ) {
    this.projection = structuredClone(projection);
  }
  get data(): ContextProjection {
    return structuredClone(this.projection);
  }
  get consumed(): boolean {
    return this.invalidated;
  }
  async verify(signal?: AbortSignal): Promise<void> {
    if (this.invalidated)
      throw new Error('Evidence binding was consumed by a host write. Prepare fresh context.');
    await this.workspace.verify(this.projection.snapshotId, this.source, signal);
  }
  async beforeEffect(request: ActuationRequest, signal?: AbortSignal): Promise<void> {
    if (!this.authorizeEffect(structuredClone(request)))
      throw new Error('Effect is outside the evidence binding write scope.');
    await this.verify(signal);
    signal?.throwIfAborted();
  }
  /** Invalidate at dispatch, including uncertain writes. No inferred rebasing of observed evidence. */
  dispatched(): void {
    this.invalidated = true;
  }
}
