import { HyperFormula } from 'hyperformula';

/**
 * HyperFormula's English dictionary has no `#NULL!` error type, but the model
 * and Excel do (formula semantics §3). A derived language adds the mapping so
 * an imported `#NULL!` literal parses and propagates as an error instead of
 * degrading to `#VALUE!`.
 */
export const CALCPLUS_LANGUAGE = 'calcplus-en';

let registered = false;
let nullErrorSupported = false;

export function registerCalcplusLanguage(): void {
  if (registered) return;
  registered = true;
  const base = HyperFormula.languages['enGB'];
  if (base === undefined) return;
  try {
    // The error translation set is keyed by HyperFormula's error enum; the
    // custom `NULL` key is the whole point of the derived language.
    const errors: Record<string, string> = { ...base.errors };
    errors.NULL = '#NULL!';
    HyperFormula.registerLanguage(CALCPLUS_LANGUAGE, {
      ...base,
      langCode: CALCPLUS_LANGUAGE,
      errors: errors as typeof base.errors,
    });
    nullErrorSupported = true;
  } catch {
    // A hot reload can find the language already registered; use it then.
    try {
      HyperFormula.getLanguage(CALCPLUS_LANGUAGE);
      nullErrorSupported = true;
    } catch {
      nullErrorSupported = false;
    }
  }
}

export function engineLanguage(): string | undefined {
  return nullErrorSupported ? CALCPLUS_LANGUAGE : undefined;
}

export function supportsNullError(): boolean {
  return nullErrorSupported;
}
