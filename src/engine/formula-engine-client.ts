import type { EngineBatchResult, EngineCellInput, EngineEdit, EngineWorkbookInput } from './port';
import type { FromEngineWorkerMessage, ToEngineWorkerMessage } from './protocol';

export type EngineWorkerFactory = () => Worker;

export interface FormulaEngineClientCallbacks {
  /**
   * `baseEditVersion` is the document edit version the batch belongs to.
   *
   * Returning a promise holds back the next request until it settles. The
   * calculation controller uses this to apply a large fan-out result in
   * bounded slices: the worker stays idle while the main thread finishes, so
   * pending results never pile up (performance §3.2).
   */
  onResult(
    result: EngineBatchResult,
    baseEditVersion: number,
    phase: 'load' | 'apply',
  ): void | Promise<void>;
  onError(message: string): void;
}

export interface FormulaEngineClientOptions {
  readonly workerFactory?: EngineWorkerFactory;
  /** Cells per `load` message. Keeps each structured clone small. */
  readonly chunkSize?: number;
}

const DEFAULT_CHUNK_SIZE = 10_000;

/**
 * Worker boundary with request coalescing and backpressure.
 *
 * - A workbook snapshot is loaded one chunk at a time: the worker acks each
 *   non-final chunk and only then does the client post the next one.
 * - At most one `apply` is in flight. Edits that arrive meanwhile are
 *   coalesced per cell, so typing quickly cannot build an unbounded queue of
 *   stale writes.
 * - A result callback may return a promise; the next request waits for it, so
 *   a slow main-thread application slows the worker down instead of buffering
 *   more results.
 */
export class FormulaEngineClient {
  private worker: Worker | null = null;
  private failed = false;
  private ready = false;
  private loadInFlight = false;
  private loadRequestId = 0;
  private loadCells: readonly EngineCellInput[] | null = null;
  private loadOffset = 0;
  private applyInFlight = false;
  /** A result callback is still applying; no new request may be posted. */
  private deliveryInFlight = false;
  private nextRequestId = 1;
  private pendingEdits = new Map<string, { readonly edit: EngineEdit; readonly version: number }>();
  private pendingBaseEditVersion = 0;
  private readonly requestMeta = new Map<
    number,
    { readonly kind: 'load' | 'apply'; readonly baseEditVersion: number }
  >();
  private readonly chunkSize: number;

  constructor(
    private readonly callbacks: FormulaEngineClientCallbacks,
    options: FormulaEngineClientOptions = {},
  ) {
    this.chunkSize = Math.max(1, options.chunkSize ?? DEFAULT_CHUNK_SIZE);
    try {
      this.worker = (options.workerFactory ?? createFormulaWorker)();
      this.worker.onmessage = (event: MessageEvent<FromEngineWorkerMessage>) => {
        void this.handleMessage(event.data);
      };
      this.worker.onerror = (event) => {
        this.fail(event.message || 'The formula worker failed.');
      };
      this.worker.onmessageerror = () => {
        this.fail('The formula worker sent a message that could not be read.');
      };
    } catch (error) {
      this.fail(error instanceof Error ? error.message : String(error));
    }
  }

  get isReady(): boolean {
    return this.ready;
  }

  get hasFailed(): boolean {
    return this.failed;
  }

  /** Pending edits after coalescing: one entry per edited cell. */
  get queuedEditCount(): number {
    return this.pendingEdits.size;
  }

  /** True while a load or apply request is with the worker, or a result is applying. */
  get isBusy(): boolean {
    return this.loadInFlight || this.applyInFlight || this.deliveryInFlight;
  }

  /**
   * Build the engine from the workbook snapshot. Cells are chunked so the main
   * thread never serialises a whole large sheet in one pause, and each chunk is
   * sent only after the worker acknowledged the previous one.
   */
  initialize(
    workbook: EngineWorkbookInput,
    cells: readonly EngineCellInput[],
    baseEditVersion: number,
  ): void {
    if (this.worker === null || this.failed) return;
    this.ready = false;
    this.loadInFlight = true;
    this.loadCells = cells;
    this.loadOffset = 0;
    const requestId = this.nextRequestId;
    this.nextRequestId += 1;
    this.loadRequestId = requestId;
    this.requestMeta.set(requestId, { kind: 'load', baseEditVersion });
    this.post({ type: 'start', requestId, workbook });
    this.postNextLoadChunk();
  }

  /** Queue a delta. Calls made while a request is in flight are coalesced. */
  applyEdits(edits: readonly EngineEdit[], baseEditVersion: number): void {
    if (this.worker === null || this.failed || edits.length === 0) return;
    for (const edit of edits) {
      const key = editKey(edit);
      // Later writes to the same cell supersede earlier ones; the map keeps
      // the queue bounded by the number of distinct touched cells.
      this.pendingEdits.set(key, { edit, version: baseEditVersion });
    }
    this.pendingBaseEditVersion = Math.max(this.pendingBaseEditVersion, baseEditVersion);
    this.flush();
  }

  /**
   * Drop queued edits at or before `version`. A rebuild calls this for the
   * version its snapshot was taken at: those edits are already in the
   * snapshot, and re-sending them after a structural edit would address cells
   * at their old coordinates. Edits made while the snapshot was read carry a
   * newer version and stay queued.
   */
  discardPendingBefore(version: number): void {
    let remainingBase = 0;
    for (const [key, pending] of this.pendingEdits) {
      if (pending.version <= version) {
        this.pendingEdits.delete(key);
      } else {
        remainingBase = Math.max(remainingBase, pending.version);
      }
    }
    this.pendingBaseEditVersion = remainingBase;
  }

  dispose(): void {
    const worker = this.worker;
    this.worker = null;
    this.ready = false;
    this.failed = true;
    this.loadInFlight = false;
    this.loadCells = null;
    this.pendingEdits.clear();
    this.requestMeta.clear();
    if (worker !== null) {
      worker.terminate();
    }
  }

  private async handleMessage(message: FromEngineWorkerMessage): Promise<void> {
    if (message.type === 'error') {
      this.requestMeta.delete(message.requestId);
      this.fail(message.message);
      return;
    }

    if (message.type === 'ack') {
      if (
        message.requestId === this.loadRequestId &&
        this.loadInFlight &&
        this.loadCells !== null
      ) {
        this.postNextLoadChunk();
      }
      return;
    }

    const meta = this.requestMeta.get(message.requestId);
    this.requestMeta.delete(message.requestId);
    if (meta === undefined) return;

    if (meta.kind === 'load') {
      // A rebuild may have started a newer load while this result was in
      // flight; its chunk positions belong to the old request.
      if (message.requestId !== this.loadRequestId) return;
      this.loadInFlight = false;
      this.loadCells = null;
      this.ready = true;
    } else {
      this.applyInFlight = false;
    }

    try {
      this.deliveryInFlight = true;
      await this.callbacks.onResult(message.result, meta.baseEditVersion, meta.kind);
    } catch (error) {
      this.fail(error instanceof Error ? error.message : String(error));
      return;
    } finally {
      this.deliveryInFlight = false;
      this.flush();
    }
  }

  private postNextLoadChunk(): void {
    const cells = this.loadCells;
    if (cells === null) return;
    const requestId = this.loadRequestId;
    const end = Math.min(this.loadOffset + this.chunkSize, cells.length);
    const final = end >= cells.length;
    const chunk = cells.slice(this.loadOffset, end);
    this.loadOffset = end;
    if (final) this.loadCells = null;
    this.post({ type: 'load', requestId, cells: chunk, final });
  }

  private flush(): void {
    if (
      this.worker === null ||
      this.failed ||
      !this.ready ||
      this.loadInFlight ||
      this.applyInFlight ||
      this.deliveryInFlight ||
      this.pendingEdits.size === 0
    ) {
      return;
    }
    const edits = [...this.pendingEdits.values()].map((pending) => pending.edit);
    const baseEditVersion = this.pendingBaseEditVersion;
    this.pendingEdits.clear();
    this.pendingBaseEditVersion = 0;
    this.applyInFlight = true;
    const requestId = this.nextRequestId;
    this.nextRequestId += 1;
    this.requestMeta.set(requestId, { kind: 'apply', baseEditVersion });
    this.post({ type: 'apply', requestId, edits });
  }

  private post(message: ToEngineWorkerMessage): void {
    if (this.worker === null || this.failed) return;
    try {
      this.worker.postMessage(message);
    } catch (error) {
      this.fail(error instanceof Error ? error.message : String(error));
    }
  }

  private fail(message: string): void {
    if (this.failed) return;
    this.failed = true;
    this.ready = false;
    this.loadInFlight = false;
    this.loadCells = null;
    this.pendingEdits.clear();
    this.requestMeta.clear();
    this.callbacks.onError(message);
  }
}

function editKey(edit: EngineEdit): string {
  return `${edit.sheet}:${edit.row}:${edit.col}`;
}

function createFormulaWorker(): Worker {
  return new Worker(new URL('./formula.worker.ts', import.meta.url), { type: 'module' });
}
