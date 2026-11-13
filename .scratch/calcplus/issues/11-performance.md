# 11 — Performance targets and benchmarks

Status: open
Type: grilling
Blocked by: none

## Question

What performance must v1 hit, and how is it measured?

Set measurable targets for:

- Max workbook size and used range the app must stay usable with
- Cold open time for a representative workbook
- Full and incremental recalculation latency
- Scroll/paint frame budget and interaction latency
- Large paste, fill, copy, and find operations
- Memory ceiling in a browser tab
- Desktop-browser-only assumption (state the cut explicitly)

Define the benchmark corpus (synthetic and real-world shapes), the measurement harness, and the CI regression gate.
