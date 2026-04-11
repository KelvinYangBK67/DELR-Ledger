# DELR Ledger Changelog

[繁體中文](CHANGELOG_zh.md)

## v0.1.1

Released as the first patch update after `v0.1.0`. This release focuses on correctness, packaging reliability, and export/import stability.

### Fixed

- Corrected ledger totals to use the stored signed amount as an algebraic sum.
- Fixed Windows packaging behavior where PyInstaller cleanup could fail with `PermissionError`.
- Fixed build scripts so failed PyInstaller runs are no longer reported as successful builds.
- Improved PDF export font handling for Chinese, Chinese punctuation, and mixed Chinese/Latin text.
- Fixed clipboard import parsing so each strongly detected field is matched only once per row.
- Improved amount display consistency, including German decimal comma formatting.

### Changed

- Build scripts now remove previous `build` and `dist` outputs before invoking PyInstaller.
- Document export and table totals now follow the same stored amount semantics.

## v0.1.0

Initial public release of DELR Ledger.

### Added

- Tkinter desktop GUI for personal ledger tracking.
- Local `.delr` ledger files with CSV-compatible content.
- Multi-language interface: Traditional Chinese, English, and German.
- Ledger creation, opening, import, and export workflows.
- Clipboard import with smart recognition for dates, entry types, amounts, units, and payment hints.
- Data import/export support for `.delr`, `.csv`, `.tsv`, `.xlsx`, `.json`, `.xml`, and `.yaml` / `.yml`.
- Document export support for Markdown, Word, and PDF.
- Table sorting, filtering, pagination by all/year/month/day, and multi-select deletion/editing.
- Per-currency totals for total, income, and expense.
- Windows build and release scripts.
