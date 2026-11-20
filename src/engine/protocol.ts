import type { EngineBatchResult, EngineCellInput, EngineEdit, EngineWorkbookInput } from './port';

/**
 * Worker protocol. Everything here is structured-clone safe; the engine's
 * rewritten formula text never appears in a message.
 */
export type ToEngineWorkerMessage =
  | {
      readonly type: 'start';
      readonly requestId: number;
      readonly workbook: EngineWorkbookInput;
    }
  | {
      readonly type: 'load';
      readonly requestId: number;
      readonly cells: readonly EngineCellInput[];
      /** The last chunk also asks the worker to finish evaluation. */
      readonly final: boolean;
    }
  | {
      readonly type: 'apply';
      readonly requestId: number;
      readonly edits: readonly EngineEdit[];
    }
  | { readonly type: 'dispose' };

export type FromEngineWorkerMessage =
  | {
      readonly type: 'result';
      readonly requestId: number;
      readonly result: EngineBatchResult;
    }
  | {
      /**
       * One non-final `load` chunk was accepted. The client sends the next
       * chunk only after this ack, so a large snapshot never has every chunk
       * cloned into the worker queue at once (performance §3.2: backpressure).
       */
      readonly type: 'ack';
      readonly requestId: number;
    }
  | {
      readonly type: 'error';
      readonly requestId: number;
      readonly message: string;
    };
