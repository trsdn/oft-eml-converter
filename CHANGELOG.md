# Changelog

All notable changes to this project are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `AGENTS.md` with the authoritative build, run, and validation commands, and the operations agents must not perform.
- Architecture notes and decision records under `docs/`.
- This changelog.
- Dependabot configuration for pip and GitHub Actions.
- Declared primary language, data handling, network behaviour, and accessibility properties in the README.
- Python 3.13 and 3.14 to the CI matrix. Both are maintained upstream, so under the support policy stated below they are tested rather than merely tolerated. The matrix now covers every maintained interpreter: 3.10 through 3.14.

### Changed

- `AGENTS.md` replaced. It previously held a development retrospective rather than operating instructions; the technical substance moved into `docs/decisions/`.
- The `pytest` floor is now 9.1.1, matching the version the supported interpreters resolve to.

### Removed

- Python 3.8 and 3.9 support. Both have reached end of life — 3.8 in October 2024, 3.9 in October 2025 — and neither receives security fixes. The supported range is now 3.10 through 3.14, in the CI matrix and the README alike. The project's policy is to test every interpreter that is still maintained upstream, and only those.

### Fixed

- The coverage upload ran only on a leg selected by comparing against a hard-coded Python version. Had that version left the test matrix, the condition would have matched no leg and coverage would have stopped being uploaded silently. The leg is now marked in the matrix itself, which Actions appends as a combination if it ever stops matching, so the upload cannot vanish unnoticed.
- The coverage upload passed `file` to `codecov-action`, an input the action stopped accepting at v4. It was silently discarded, leaving the upload to whatever its own file search happened to find. The input is now `files`, so the report that is uploaded is again the one the step names.
- No pull request could be merged, however green its checks. The branch ruleset requires a status check named `integration-test`, but the job of that name was a matrix, and a matrix reports one check per leg with the matrix values appended — so the bare name was never reported. The matrix job is now `integration`, and a dedicated `integration-test` job gates on its result. That job fails on any result other than success, because a `needs` dependency on a skipped job otherwise counts as satisfied, which would have let the requirement pass exactly when the integration tests had not run.

## [1.1.0] — 2025-09-05

### Added

- Cross-platform CI matrix covering Windows, macOS, and Linux on Python 3.8 through 3.12.
- Validation of CLI converter error paths.
- Virtual environment creation and management tests.
- Headless-safe GUI import test, so the suite runs on CI runners without a display.
- Cross-platform file operation tests.
- Package requirement verification.
- Release automation through GitHub Actions.

### Fixed

- Unicode encoding failures in the Windows console.

## [1.0.0] — 2025-09-05

### Added

- Initial release.
- Conversion of `.oft` templates to RFC-compliant `.eml` messages.
- Inline image support with Content-ID mapping preserved.
- Batch processing.
- Tkinter GUI and command-line interface.
- Cross-platform support.

[Unreleased]: https://github.com/trsdn/oft-eml-converter/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/trsdn/oft-eml-converter/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/trsdn/oft-eml-converter/releases/tag/v1.0.0
