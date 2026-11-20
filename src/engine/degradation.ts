/**
 * The degradation ladder (performance §8).
 *
 * When a workbook outgrows the guaranteed envelope, calcplus degrades visibly
 * and in a fixed order: the large-workbook banner, conditional formatting
 * limited to the visible range, less scroll overscan, volatile recalculation
 * moved to idle time, a memory warning at 1 GB and a refusal of new large
 * operations at 1.2 GB. Rendering degrades before correctness does: this module
 * never reports that a value may be dropped, and it never edits the model.
 *
 * The function is pure so the ladder can be tested at every boundary. Wiring
 * (showing the banner, limiting the paint or pausing volatile formulas) reads
 * the plan; it does not decide anything itself.
 */

export const GUARANTEED_TOUCHED_CELLS = 250_000;
export const GUARANTEED_FORMULAS = 50_000;
export const DEGRADED_TOUCHED_CELLS = 1_000_000;
export const DEGRADED_FORMULAS = 200_000;

/** The guaranteed envelope must stay under 500 MB (performance §4). */
export const GUARANTEED_WORKING_SET_BYTES = 500 * 1024 ** 2;
/** At 1 GB a visible memory warning appears (performance §4). */
export const MEMORY_WARNING_BYTES = 1 * 1024 ** 3;
/** At 1.2 GB new large operations are refused (performance §4). */
export const MEMORY_REFUSE_BYTES = Math.floor(1.2 * 1024 ** 3);

export type SizeEnvelope = 'guaranteed' | 'degraded' | 'refused';
export type WorkloadPhase = 'open' | 'runtime';

export interface DegradationInput {
  readonly touchedCells: number;
  readonly formulas: number;
  /** The working-set estimate, when the caller has one. */
  readonly estimatedWorkingSetBytes?: number;
  /** An open may refuse an oversized file; a running workbook only refuses work. */
  readonly phase?: WorkloadPhase;
}

export interface DegradationBanner {
  readonly level: 'info' | 'warning' | 'critical';
  readonly message: string;
}

export interface DegradationPlan {
  readonly phase: WorkloadPhase;
  readonly envelope: SizeEnvelope;
  readonly touchedCells: number;
  readonly formulas: number;
  readonly estimatedWorkingSetBytes: number | null;
  /** The most severe notice, ready for the banner. */
  readonly banner: DegradationBanner | null;
  /** Every notice in ladder order, for a report or the About panel. */
  readonly notices: readonly string[];
  /** Step 2: conditional formatting paints only the visible range. */
  readonly conditionalFormatScope: 'full' | 'visible';
  /** Step 3: the grid keeps less overscan around the viewport. */
  readonly scrollOverscan: 'full' | 'reduced';
  /** Step 4: volatile formulas (TODAY, NOW) wait for idle time. */
  readonly volatileRecalculation: 'immediate' | 'idle';
  /** Step 5: a visible memory warning. */
  readonly memoryWarning: boolean;
  /** Step 6: new large operations ask first and are then refused. */
  readonly refuseLargeOperations: boolean;
  readonly promptToClose: boolean;
  /** The workbook is above the refusal line and was not imported or modified. */
  readonly refusesWorkbook: boolean;
  /** An estimate above 1.2 GB at open refuses the open. */
  readonly refusesOpen: boolean;
  /**
   * Always true. The ladder slows rendering down; it never drops or rewrites a
   * value, and it never touches the file on disk.
   */
  readonly valuesPreserved: true;
}

/**
 * Decide the ladder step for a workbook. `touchedCells` and `formulas` are the
 * authored counts, not the grid's dimensions.
 */
export function degradationPlan(input: DegradationInput): DegradationPlan {
  const touchedCells = Math.max(0, Math.floor(input.touchedCells));
  const formulas = Math.max(0, Math.floor(input.formulas));
  const estimatedWorkingSetBytes =
    input.estimatedWorkingSetBytes === undefined
      ? null
      : Math.max(0, input.estimatedWorkingSetBytes);
  const phase = input.phase ?? 'runtime';

  const envelope = classifyEnvelope(touchedCells, formulas);
  const refusesWorkbook = envelope === 'refused';
  const overGuaranteed = envelope !== 'guaranteed';
  const memoryWarning =
    estimatedWorkingSetBytes !== null && estimatedWorkingSetBytes >= MEMORY_WARNING_BYTES;
  const memoryRefusal =
    estimatedWorkingSetBytes !== null && estimatedWorkingSetBytes >= MEMORY_REFUSE_BYTES;
  const refusesOpen = refusesWorkbook || (phase === 'open' && memoryRefusal);
  const refuseLargeOperations = memoryRefusal;

  const notices: string[] = [];
  let banner: DegradationBanner | null = null;

  const raise = (next: DegradationBanner): void => {
    notices.push(next.message);
    if (banner === null || rank(next.level) > rank(banner.level)) banner = next;
  };

  if (refusesWorkbook) {
    raise({ level: 'critical', message: refuseMessage(touchedCells, formulas) });
  } else if (overGuaranteed) {
    raise({
      level: 'info',
      message:
        'Large workbook: calculation runs when the app is idle and conditional formatting is limited to the visible cells. Values are complete.',
    });
  }

  if (memoryRefusal) {
    raise({
      level: 'critical',
      message:
        'Memory is above 1.2 GB. New large operations are refused; close the workbook to free memory.',
    });
  } else if (memoryWarning) {
    raise({
      level: 'warning',
      message: 'Memory is above 1 GB. Large operations will ask before they run.',
    });
  }

  const promptToClose = memoryRefusal;

  return {
    phase,
    envelope,
    touchedCells,
    formulas,
    estimatedWorkingSetBytes,
    banner,
    notices,
    // Step 2 only engages past the guaranteed envelope.
    conditionalFormatScope: overGuaranteed ? 'visible' : 'full',
    // Step 3 engages with step 2; the two are one visual step.
    scrollOverscan: overGuaranteed ? 'reduced' : 'full',
    // Step 4 moves volatile recalculation to idle time, never the whole engine.
    volatileRecalculation: overGuaranteed ? 'idle' : 'immediate',
    memoryWarning,
    refuseLargeOperations,
    promptToClose,
    refusesWorkbook,
    refusesOpen,
    valuesPreserved: true,
  };
}

export function classifyEnvelope(touchedCells: number, formulas: number): SizeEnvelope {
  if (touchedCells > DEGRADED_TOUCHED_CELLS || formulas > DEGRADED_FORMULAS) return 'refused';
  if (touchedCells > GUARANTEED_TOUCHED_CELLS || formulas > GUARANTEED_FORMULAS) return 'degraded';
  return 'guaranteed';
}

/**
 * True when a large operation should ask before running: at 1 GB the ladder
 * warns, at 1.2 GB it refuses. Everything below 1 GB is the guaranteed or
 * degraded envelope and needs no prompt.
 */
export function shouldConfirmLargeOperation(plan: DegradationPlan): boolean {
  return plan.memoryWarning;
}

/** True when the ladder refuses a new large operation (1.2 GB). */
export function shouldRefuseLargeOperation(plan: DegradationPlan): boolean {
  return plan.refuseLargeOperations;
}

export function describeDegradation(plan: DegradationPlan): string {
  if (plan.refusesWorkbook)
    return `${plan.banner?.message ?? 'The workbook was refused.'} Nothing was imported or changed.`;
  const parts = [`envelope ${plan.envelope}`];
  if (plan.conditionalFormatScope === 'visible') parts.push('conditional formatting visible-only');
  if (plan.scrollOverscan === 'reduced') parts.push('reduced overscan');
  if (plan.volatileRecalculation === 'idle') parts.push('volatile recalculation on idle');
  if (plan.memoryWarning) parts.push('memory warning');
  if (plan.refuseLargeOperations) parts.push('large operations refused');
  parts.push('values preserved');
  return parts.join(', ');
}

function refuseMessage(touchedCells: number, formulas: number): string {
  const limits = `${DEGRADED_TOUCHED_CELLS.toLocaleString('en-US')} touched cells or ${DEGRADED_FORMULAS.toLocaleString('en-US')} formulas`;
  return `This workbook has ${touchedCells.toLocaleString('en-US')} touched cells and ${formulas.toLocaleString('en-US')} formulas; calcplus supports up to ${limits}. The file was not modified.`;
}

function rank(level: DegradationBanner['level']): number {
  return level === 'critical' ? 2 : level === 'warning' ? 1 : 0;
}
