# 13 — Testing and Excel-conformance strategy

Status: open
Type: grilling
Blocked by: 02, 03, 04, 18, 19, 20, 21, 22, 24

## Question

How do we prove calcplus is faithful?

Define:

- Formula conformance: golden workbooks, generated cases, comparison against real Excel, tolerance rules
- `.xlsx` round-trip fixtures and diffing strategy
- Grid UI/interaction regression testing
- Performance benchmarks in CI
- Unit/integration boundaries around engine, model, and view
- The minimum bar before v1 can be called "faithful", and what is explicitly not tested

Depends on the semantics contract (02), the engine choice (03), and the I/O approach (04).
