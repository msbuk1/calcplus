import type { EngineBatchResult, EngineCellInput, EngineEdit, EngineWorkbookInput } from './port';
import type { FromEngineWorkerMessage, ToEngineWorkerMessage } from './protocol';

export type EngineWorkerFactory = () => Worker;

export interface FormulaEngineClientCallbacks {
  /** `baseEditVersion` is the document edit version the batch belongs to. */
  onResult(result: EngineBatchResult, baseEditVersion: number, phase: 'load' | 'apply'): void;
  onError(message: string): void;
}

export interface FormulaEngineClientOptions {
  readonly workerFactory?: EngineWorkerFactory;
  /** Cells per `load` message. Keeps each structured clone small. */
  readonly chunkSize?: number;
}

const DEFAULT_CHUNK_SIZE = 10_000;

/** Worker boundary with request coalescing and backpressure. */
export class FormulaEngineClient {
  private worker: Worker | null = null;
  private failed = false;
  private ready = false;
  private loadInFlight = false;
  private applyInFlight = false;
  private nextRequestId = 1;
  private pendingEdits: EngineEdit[] = [];
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
        this.handleMessage(event.data);
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

  get queuedEditCount(): number {
    return this.pendingEdits.length;
  }

  /**
   * Build the engine from the workbook snapshot. Cells are chunked so the
   * main thread never serialises a whole large sheet in one pause.
   */
  initialize(
    workbook: EngineWorkbookInput,
    cells: readonly EngineCellInput[],
    baseEditVersion: number,
  ): void {
    if (this.worker === null || this.failed) return;
    this.ready = false;
    this.loadInFlight = true;
    const requestId = this.nextRequestId;
    this.nextRequestId += 1;
    this.requestMeta.set(requestId, { kind: 'load', baseEditVersion });
    this.post({ type: 'start', requestId, workbook });

    if (cells.length === 0) {
      this.post({ type: 'load', requestId, cells: [], final: true });
      return;
    }
    for (let index = 0; index < cells.length; index += this.chunkSize) {
      const final = index + this.chunkSize >= cells.length;
      this.post({
        type: 'load',
        requestId,
        cells: cells.slice(index, index + this.chunkSize),
        final,
      });
    }
  }

  /** Queue a delta. Calls made while a request is in flight are coalesced. */
  applyEdits(edits: readonly EngineEdit[], baseEditVersion: number): void {
    if (this.worker === null || this.failed || edits.length === 0) return;
    this.pendingEdits.push(...edits);
    this.pendingBaseEditVersion = Math.max(this.pendingBaseEditVersion, baseEditVersion);
    this.flush();
  }

  dispose(): void {
    const worker = this.worker;
    this.worker = null;
    this.ready = false;
    this.failed = true;
    this.pendingEdits = [];
    this.requestMeta.clear();
    if (worker !== null) {
      worker.terminate();
    }
  }

  private handleMessage(message: FromEngineWorkerMessage): void {
    if (message.type === 'error') {
      this.requestMeta.delete(message.requestId);
      this.fail(message.message);
      return;
    }

    const meta = this.requestMeta.get(message.requestId);
    this.requestMeta.delete(message.requestId);
    if (meta === undefined) return;

    if (meta.kind === 'load') {
      this.loadInFlight = false;
      this.ready = true;
    } else {
      this.applyInFlight = false;
    }
    try {
      this.callbacks.onResult(message.result, meta.baseEditVersion, meta.kind);
    } finally {
      this.flush();
    }
  }

  private flush(): void {
    if (
      this.worker === null ||
      this.failed ||
      !this.ready ||
      this.loadInFlight ||
      this.applyInFlight ||
      this.pendingEdits.length === 0
    ) {
      return;
    }
    const edits = this.pendingEdits;
    const baseEditVersion = this.pendingBaseEditVersion;
    this.pendingEdits = [];
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
    this.pendingEdits = [];
    this.requestMeta.clear();
    this.callbacks.onError(message);
  }
}

function createFormulaWorker(): Worker {
  return new Worker(new URL('./formula.worker.ts', import.meta.url), { type: 'module' });
}
