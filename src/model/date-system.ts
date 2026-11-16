/**
 * Excel's two date systems. Dates are serial numbers; the workbook flag only
 * decides how the serial is interpreted (formula semantics §5).
 *
 * 1900 system:
 * - serial 1 is 1900-01-01 and serial 59 is 1900-02-28.
 * - serial 60 is Excel's fake 29 February 1900. Excel reports 1900 as a leap
 *   year although it was not, so every serial from 61 onwards is one day ahead
 *   of the Gregorian calendar.
 * - serial 0 renders as 1 January 1900 day zero, as Excel does.
 *
 * 1904 system: serial 0 is 1904-01-01 and leap days follow the Gregorian
 * calendar.
 */

export const FAKE_1900_LEAP_DAY_SERIAL = 60;

const MS_PER_DAY = 86_400_000;

export interface DateParts {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hours: number;
  readonly minutes: number;
  readonly seconds: number;
}

/** True when the serial is Excel's invented 29 February 1900. */
export function isFake1900LeapDay(serial: number): boolean {
  return Math.floor(serial) === FAKE_1900_LEAP_DAY_SERIAL;
}

export function workbookDateSystem(date1904: boolean): '1900' | '1904' {
  return date1904 ? '1904' : '1900';
}

/**
 * Interpret a serial number as a calendar date and time. The result can carry
 * day 0 for the 1900 system's serial zero. `null` means the serial has no
 * Excel date representation in that system.
 */
export function serialToDateParts(serial: number, date1904: boolean): DateParts | null {
  if (!Number.isFinite(serial)) return null;
  const day = Math.floor(serial);
  const fraction = serial - day;
  const milliseconds = Math.round(fraction * MS_PER_DAY);

  if (date1904) {
    if (day < 0) return null;
    const date = new Date(Date.UTC(1904, 0, 1) + day * MS_PER_DAY + milliseconds);
    return partsFromUtc(date);
  }

  if (day < 0) return null;
  if (day === 0) {
    return {
      year: 1900,
      month: 1,
      day: 0,
      hours: Math.floor(milliseconds / 3_600_000),
      minutes: Math.floor((milliseconds % 3_600_000) / 60_000),
      seconds: Math.floor((milliseconds % 60_000) / 1000),
    };
  }
  if (day === FAKE_1900_LEAP_DAY_SERIAL) {
    return {
      year: 1900,
      month: 2,
      day: 29,
      hours: Math.floor(milliseconds / 3_600_000),
      minutes: Math.floor((milliseconds % 3_600_000) / 60_000),
      seconds: Math.floor((milliseconds % 60_000) / 1000),
    };
  }
  const epochDay = day < FAKE_1900_LEAP_DAY_SERIAL ? day : day - 1;
  const date = new Date(Date.UTC(1900, 0, 1) + (epochDay - 1) * MS_PER_DAY + milliseconds);
  return partsFromUtc(date);
}

/** Serial number for a calendar date and time, reproducing serial 60. */
export function datePartsToSerial(
  year: number,
  month: number,
  day: number,
  date1904: boolean,
  hours = 0,
  minutes = 0,
  seconds = 0,
): number {
  // Excel's DATE(1900,2,29) is the invented leap day, before Gregorian
  // normalisation would turn it into 1 March.
  if (!date1904 && year === 1900 && month === 2 && day === 29) {
    return FAKE_1900_LEAP_DAY_SERIAL + hours / 24 + minutes / 1440 + seconds / 86_400;
  }
  const normalized = normalizeDateParts(year, month, day);
  if (date1904) {
    const time = Date.UTC(
      normalized.year,
      normalized.month - 1,
      normalized.day,
      hours,
      minutes,
      seconds,
    );
    return (time - Date.UTC(1904, 0, 1)) / MS_PER_DAY;
  }
  const time = Date.UTC(
    normalized.year,
    normalized.month - 1,
    normalized.day,
    hours,
    minutes,
    seconds,
  );
  const days = (time - Date.UTC(1899, 11, 31)) / MS_PER_DAY;
  // Serials from 1 March 1900 onwards include Excel's phantom 29 February.
  return normalized.year > 1900 || (normalized.year === 1900 && normalized.month >= 3)
    ? days + 1
    : days;
}

/** Whole-day serial for a `Date`, interpreted in the workbook's date system. */
export function dateToSerial(date: Date, date1904: boolean): number {
  return datePartsToSerial(
    date.getUTCFullYear(),
    date.getUTCMonth() + 1,
    date.getUTCDate(),
    date1904,
    date.getUTCHours(),
    date.getUTCMinutes(),
    date.getUTCSeconds(),
  );
}

/** Excel's DATE() semantics for an out-of-range month or day. */
export function normalizeDateParts(
  year: number,
  month: number,
  day: number,
): { year: number; month: number; day: number } {
  const base = new Date(Date.UTC(year, month - 1, 1));
  const shifted = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth(), day));
  return {
    year: shifted.getUTCFullYear(),
    month: shifted.getUTCMonth() + 1,
    day: shifted.getUTCDate(),
  };
}

/** True when a serial is in the window affected by Excel's 1900 leap bug. */
export function isBefore1900March(serial: number): boolean {
  return serial < 61;
}

function partsFromUtc(date: Date): DateParts {
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
  };
}
