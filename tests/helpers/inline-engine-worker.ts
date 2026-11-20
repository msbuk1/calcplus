import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import type { FromEngineWorkerMessage, ToEngineWorkerMessage } from '@engine/protocol';

/**
 * An inline stand-in for the formula Web Worker.
 *
 * `CalculationController` owns a Worker, so a Node test cannot use the real
 * one (the comment in `tests/m6-roundtrip.test.ts` records the same
 * constraint). This helper speaks the worker protocol and runs the real
 * `HyperFormulaEngine` adapter, so the queue, the load chunks, the acks and the
 * fan-out results are the production shapes. Messages run one at a time, like
 * the real worker's event loop.
 *
 * `autoRun: false` leaves messages in `inbox` so a test can inspect the
 * backpressure (only one load chunk may be outstanding) before pumping it.
 */
export class InlineEngineWorker {
  onmessage: ((event: MessageEvent<FromEngineWorkerMessage>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: (() => void) | null = null;

  readonly inbox: ToEngineWorkerMessage[] = [];
  /** Every message the worker received, in order. */
  readonly received: ToEngineWorkerMessage[] = [];
  /** Every load chunk the engine received, in order. */
  readonly loadChunks: { readonly cells: number; readonly final: boolean }[] = [];
  private readonly engine = new HyperFormulaEngine();
  private draining = false;
  private terminated = false;

  constructor(private readonly autoRun = true) {}

  postMessage(message: ToEngineWorkerMessage): void {
    if (this.terminated) return;
    this.inbox.push(message);
    this.received.push(message);
    if (this.autoRun) void this.drain();
  }

  terminate(): void {
    this.terminated = true;
    this.inbox.length = 0;
    this.engine.dispose();
  }

  /** Run one queued message. Only useful with `autoRun: false`. */
  async step(): Promise<void> {
    const message = this.inbox.shift();
    if (message !== undefined) await this.handle(message);
  }

  /** Run everything queued, following any message the client posts in reply. */
  async drain(): Promise<void> {
    if (this.draining) return;
    this.draining = true;
    try {
      while (this.inbox.length > 0) {
        const message = this.inbox.shift()!;
        await this.handle(message);
      }
    } finally {
      this.draining = false;
    }
  }

  /** Wait until the worker and the client have stopped exchanging messages. */
  async settle(timeoutMs = 5_000): Promise<void> {
    const deadline = Date.now() + timeoutMs;
    for (;;) {
      if (this.inbox.length > 0) await this.drain();
      // Let the client's microtasks publish the next request or result.
      await new Promise((resolve) => {
        setTimeout(resolve, 0);
      });
      if (this.inbox.length === 0 && !this.draining) return;
      if (Date.now() > deadline) throw new Error('The inline engine worker did not settle');
    }
  }

  private async handle(message: ToEngineWorkerMessage): Promise<void> {
    await Promise.resolve();
    try {
      switch (message.type) {
        case 'start':
          this.engine.start(message.workbook);
          return;
        case 'load':
          this.loadChunks.push({ cells: message.cells.length, final: message.final });
          this.engine.load(message.cells);
          if (message.final) {
            this.emit({
              type: 'result',
              requestId: message.requestId,
              result: this.engine.finishLoad(),
            });
          } else {
            this.emit({ type: 'ack', requestId: message.requestId });
          }
          return;
        case 'apply':
          this.emit({
            type: 'result',
            requestId: message.requestId,
            result: this.engine.apply(message.edits),
          });
          return;
        case 'dispose':
          this.engine.dispose();
          return;
      }
    } catch (error) {
      this.emit({
        type: 'error',
        requestId: message.type === 'dispose' ? 0 : message.requestId,
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }

  private emit(message: FromEngineWorkerMessage): void {
    this.onmessage?.({ data: message } as MessageEvent<FromEngineWorkerMessage>);
  }
}
