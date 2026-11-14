# Contributing to calcplus

calcplus is licensed under **GPL-3.0-only**. By contributing, you agree that
your contribution is licensed under the same terms.

## Developer Certificate of Origin

Every commit must be signed off. Use:

```bash
git commit -s -m "Your message"
```

The sign-off certifies the [Developer Certificate of Origin 1.1](https://developercertificate.org/):

```text
Developer Certificate of Origin
Version 1.1

Copyright (C) 2004, 2006 The Linux Foundation and its contributors.

Everyone is permitted to copy and distribute verbatim copies of this
license document, but changing it is not allowed.

Developer's Certificate of Origin 1.1

By making a contribution to this project, I certify that:

(a) The contribution was created in whole or in part by me and I
    have the right to submit it under the open source license
    indicated in the file; or

(b) The contribution is based upon previous work that, to the best
    of my knowledge, is covered under an appropriate open source
    license and I have the right under that license to submit that
    work with modifications, whether created in whole or in part
    by me, under the same open source license (unless I am
    permitted to submit under a different license), as indicated
    in the file; or

(c) The contribution was provided directly to me by some other
    person who certified (a), (b) or (c) and I have not modified
    it.

(d) I understand and agree that this project and the contribution
    are public and that a record of the contribution (including all
    personal information I submit with it, including my sign-off) is
    maintained indefinitely and may be redistributed consistent with
    this project or the open source license(s) involved.
```

There is **no contributor licence agreement (CLA)**. The DCO is enough.

## Third-party code and assets

- A new dependency must be compatible with GPL-3.0-only. Check it before you
  add it and update `THIRD_PARTY_NOTICES.md`.
- Do not copy code from GenOffice's `ee/` directory; it is proprietary.
- Apache-2.0, MIT, ISC, BSD, MPL-2.0 and OFL-1.1 code is compatible, but keep
  the notices and attribution.
- Do not copy Microsoft icons, logos, fonts (other than the open-source
  Selawik), or other Microsoft assets.
- Test workbooks must be anonymised and cleared for test use.

## Development rules

- The specification is the source of truth: start at
  [`.scratch/calcplus/map.md`](.scratch/calcplus/map.md).
- Follow the contracts for the workbook model, commands, file compatibility,
  performance, accessibility and conformance.
- Add tests with the change. Conformance is the release gate, not coverage.
- Run the local checks before committing and keep them passing: `npm run check`,
  `npm run format:check`, `npm run build` and `npm run test:e2e`. No test or
  check runs on GitHub.
- Keep the working tree clean and the main branch passing the local checks.

## Review

- One maintainer approval is required.
- Licence, security and data-loss changes need an explicit review note.
- Report security issues privately to the maintainer rather than in a public
  issue.
