/**
 * Formula engine worker. HyperFormula is imported here and nowhere on the
 * main thread, so engine build and recalculation stay off the UI thread
 * (technical design §6).
 */
import type { EngineBatchResult } from './port';
import { HyperFormulaEngine } from './hyperformula-engine';
import type { FromEngineWorkerMessage, ToEngineWorkerMessage } from './protocol';

interface WorkerScope {
  addEventListener(
    type: 'message',
    listener: (event: MessageEvent<ToEngineWorkerMessage>) => void,
  ): void;
  postMessage(message: FromEngineWorkerMessage): void;
}

const scope = self as unknown as WorkerScope;
const engine = new HyperFormulaEngine();

scope.addEventListener('message', (event) => {
  const message = event.data;
  if (message.type === 'dispose') {
    engine.dispose();
    return;
  }
  try {
    switch (message.type) {
      case 'start':
        engine.start(message.workbook);
        return;
      case 'load':
        engine.load(message.cells);
        if (message.final) {
          respond(message.requestId, engine.finishLoad());
        } else {
          // Ask for the next chunk only once this one is accepted. A large
          // snapshot therefore stays one chunk deep in the worker queue.
          scope.postMessage({ type: 'ack', requestId: message.requestId });
        }
        return;
      case 'apply':
        respond(message.requestId, engine.apply(message.edits));
        return;
    }
  } catch (error) {
    scope.postMessage({
      type: 'error',
      requestId: message.requestId,
      message: error instanceof Error ? error.message : String(error),
    });
  }
});

function respond(requestId: number, result: EngineBatchResult): void {
  scope.postMessage({ type: 'result', requestId, result });
}
