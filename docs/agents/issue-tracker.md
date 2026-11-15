# Project records

Use these paths only when the task asks for a research note, decision, or build-status record. Do not create tickets as routine documentation.

- Product decisions and resolved specification questions: `.scratch/calcplus/`; the map is `map.md`, the contracts are in `spec/`, and decision records are in `issues/`.
- Implementation and milestone status: `.scratch/calcplus-build/issues/`.
- Research notes: `.scratch/calcplus/research/`.

Follow the format of nearby files and preserve their status fields. Read the referenced ticket when the user provides a path or number. Do not reopen a resolved decision ticket unless asked.

- For a new decision ticket or build-status record, use the next unused number in that directory. Record research notes with a descriptive file name.
- Record a decision under `## Answer`; update the map only when an approved product decision changes.
- Preserve milestone gate evidence. Mark a milestone accepted only after every exit criterion passes.
- Each milestone uses a light sense check only. Deeper hardening runs only when the user explicitly kicks it off, and nothing automatic starts it.
