# V2-F-01 — Tier 1 function tranche

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 5–9 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none (soft: V2-F-16)
- **Contract:** [formula semantics](../../spec/formula-semantics.md) §12

## Scope

Add the 37 functions recorded in formula semantics §12, taking the published
set from 52 to 89:

- Conditional aggregates: SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS,
  MINIFS
- Text: TEXT, FIND, SEARCH, SUBSTITUTE, REPLACE, TEXTJOIN, CONCAT, PROPER,
  CLEAN, VALUE
- Logic: IFS, SWITCH
- Lookup: XLOOKUP, LOOKUP
- Analysis: SUMPRODUCT, LARGE, SMALL, RANK.EQ, SUBTOTAL
- Dates: DAYS, DATEDIF, EDATE, EOMONTH, NETWORKDAYS, WEEKDAY, WEEKNUM
- Maths: TRUNC, COLUMNS, MROUND, CEILING, FLOOR

## Acceptance

- [ ] One Excel golden per function passes within the v1 tolerance.
- [ ] Error propagation, coercion and wildcards match Excel for each family.
- [ ] The registry gate and Insert Function catalogue both report 89 functions.
- [ ] Argument-count checks cover the new set; a bad call is refused like
      Excel.
- [ ] Formula text and cached values round-trip; Excel 365 opens the file
      without repair.

## Notes

`RANK.EQ` ships with the legacy `RANK` alias. `TEXT` reuses the number-format
engine. `SUMPRODUCT` needs array evaluation inside the function. `XLOOKUP`'s
scalar form can ship before V2-S-02; its array-return form waits for spill and
is out of this ticket's acceptance.
