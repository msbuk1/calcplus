import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { createSheetProjection } from '@view/projection';
import {
  ALLOWED_LINK_SCHEMES,
  checkExternalLink,
  internalLocationText,
  parseInternalLocation,
} from '@view/hyperlinks';

/**
 * The interaction-layer follow-ups: the scheme allowlist, internal `location`
 * parsing and the projection seams the grid reads for link styling, hover
 * tooltips and comment popups (feature matrix §8).
 */

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

function projectionOf(name: string) {
  const opened = openWorkbook(fixture(name));
  return {
    opened,
    projection: createSheetProjection(opened.workbook, opened.initialSheetId, { version: 1 }),
  };
}

describe('external hyperlink scheme allowlist', () => {
  it('allows http, https and mailto only', () => {
    expect([...ALLOWED_LINK_SCHEMES].sort()).toEqual(['http:', 'https:', 'mailto:']);
    for (const target of [
      'http://example.invalid/report',
      'https://example.invalid/report',
      'mailto:ana@example.invalid',
      'HTTPS://EXAMPLE.INVALID/Report',
    ]) {
      const check = checkExternalLink(target);
      expect(check.kind, target).toBe('open');
      if (check.kind === 'open') expect(check.url).toContain('example.invalid');
    }
  });

  it('normalises the scheme the browser would use', () => {
    const upper = checkExternalLink('HTTPS://Example.Invalid/Report');
    expect(upper.kind).toBe('open');
    if (upper.kind !== 'open') throw new Error('expected an allowed link');
    // The URL parser lower-cases the protocol and normalises the host, so the
    // check is against what a browser would actually do.
    expect(upper.url).toBe('https://example.invalid/Report');
  });

  it('refuses javascript, file, data, vbscript and unknown schemes', () => {
    for (const [target, scheme] of [
      ['javascript:alert(1)', 'javascript:'],
      ['JaVaScRiPt:alert(1)', 'javascript:'],
      ['file:///etc/passwd', 'file:'],
      ['data:text/html,<script>alert(1)</script>', 'data:'],
      ['vbscript:msgbox(1)', 'vbscript:'],
      ['c:\\temp\\book.xlsx', 'c:'],
    ] as const) {
      const check = checkExternalLink(target);
      expect(check.kind, target).toBe('refused');
      if (check.kind === 'refused') {
        expect(check.message).toContain('Hyperlink not followed');
        expect(check.message).toContain('only http, https and mailto');
        expect(check.message).toContain(scheme);
      }
    }
  });

  it('refuses a bare host name, a relative path and an empty target', () => {
    for (const target of ['www.example.invalid/report', '/local/report.html']) {
      const check = checkExternalLink(target);
      expect(check.kind, target).toBe('refused');
      expect(check.kind === 'refused' ? check.message : '').toContain(
        'only http, https and mailto',
      );
    }
    const empty = checkExternalLink('   ');
    expect(empty.kind).toBe('refused');
    expect(empty.kind === 'refused' ? empty.message : '').toContain('no address');
  });

  it('cannot be tricked by control characters hidden inside a scheme', () => {
    const check = checkExternalLink('java\nscript:alert(1)');
    expect(check.kind).toBe('refused');
    if (check.kind === 'refused') expect(check.message).toContain('javascript:');
  });
});

describe('internal hyperlink location parsing', () => {
  it('splits a plain sheet-qualified location', () => {
    expect(parseInternalLocation('Sheet2!A1')).toEqual({ sheetName: 'Sheet2', reference: 'A1' });
    expect(parseInternalLocation('Sheet2!A1:B4')).toEqual({
      sheetName: 'Sheet2',
      reference: 'A1:B4',
    });
  });

  it('decodes a quoted sheet name, including a doubled apostrophe', () => {
    expect(parseInternalLocation("'Target Sheet'!B3")).toEqual({
      sheetName: 'Target Sheet',
      reference: 'B3',
    });
    expect(parseInternalLocation("'Ana''s Sheet'!$A$1")).toEqual({
      sheetName: "Ana's Sheet",
      reference: '$A$1',
    });
  });

  it('splits at the separator, so a quoted name may contain an exclamation mark', () => {
    expect(parseInternalLocation("'A!B'!C3")).toEqual({ sheetName: 'A!B', reference: 'C3' });
  });

  it('treats a bare reference as a jump on the current sheet', () => {
    expect(parseInternalLocation('A1')).toEqual({ sheetName: null, reference: 'A1' });
    expect(parseInternalLocation('A1:B2')).toEqual({ sheetName: null, reference: 'A1:B2' });
  });

  it('tolerates a leading hash and surrounding whitespace', () => {
    expect(parseInternalLocation(' #Sheet1!A1 ')).toEqual({ sheetName: 'Sheet1', reference: 'A1' });
  });

  it('returns null for text that cannot be a location', () => {
    for (const location of ['', '   ', 'Sheet1!', '!A1', "'Sheet1!A1", '#']) {
      expect(parseInternalLocation(location), JSON.stringify(location)).toBeNull();
    }
  });

  it('writes the reference text back with a quoted sheet name', () => {
    expect(internalLocationText({ sheetName: 'Target Sheet', reference: 'B3:C4' })).toBe(
      "'Target Sheet'!B3:C4",
    );
    expect(internalLocationText({ sheetName: "Ana's", reference: 'A1' })).toBe("'Ana''s'!A1");
    expect(internalLocationText({ sheetName: null, reference: 'A1' })).toBe('A1');
  });
});

describe('projection hyperlink and comment seams', () => {
  it('exposes the hyperlink anchored to a cell', () => {
    const { projection } = projectionOf('validation-cf.xlsx');
    // F13 anchors an external link with a tooltip on A8 and an internal jump on
    // B8 (file-compat §9 F13).
    expect(projection.hyperlinkAt(7, 0)).toEqual({
      target: 'https://example.invalid/report',
      tooltip: 'External report',
    });
    expect(projection.hyperlinkAt(7, 1)).toEqual({ target: '', location: 'Rules!A1' });
    expect(projection.hyperlinkAt(0, 0)).toBeNull();
    // The cell record carries the link too, so the painter styles it without
    // a second lookup.
    expect(projection.cellAt(7, 0)?.hyperlink?.target).toBe('https://example.invalid/report');
    expect(projection.cellAt(0, 0)?.hyperlink).toBeNull();
  });

  it('exposes a comment record per anchor, including an empty cell', () => {
    const { projection } = projectionOf('threaded-comments.xlsx');
    expect(projection.commentAt(1, 3)).toEqual({
      text: 'Legacy note: confirm the budget.',
      author: 'Ana Silva',
      threaded: false,
    });
    expect(projection.commentAt(2, 3)).toEqual({
      text: 'Threaded: two candidates shortlisted.',
      author: 'Ana Silva',
      threaded: true,
    });
    expect(projection.commentAt(3, 3)).toEqual({
      text: 'Reply: the launch date is blocked.',
      author: 'Bo Chen',
      threaded: true,
    });
    // A comment may be anchored to a cell with no value.
    expect(projection.cellAt(1, 4)).toBeNull();
    expect(projection.commentAt(1, 4)?.text).toBe('Follow-up on an empty cell.');
    expect(projection.commentAt(0, 0)).toBeNull();
  });

  it('returns null outside the sheet bounds', () => {
    const { projection } = projectionOf('threaded-comments.xlsx');
    expect(projection.hyperlinkAt(-1, 0)).toBeNull();
    expect(projection.commentAt(0, -1)).toBeNull();
    expect(projection.commentAt(projection.rowCount, 0)).toBeNull();
  });
});
