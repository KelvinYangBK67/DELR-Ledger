# DELR Ledger Changelog

[繁體中文](CHANGELOG_zh.md)

## v0.1.2

This patch improves entry input and page navigation ergonomics.

### Added

- Added first-page, last-page, and direct page-jump controls to paginated ledger views.
- Added localized page-jump labels for Traditional Chinese, English, and German.

### Changed

- Amount fields in the add and edit forms now use the same parser as clipboard import, accepting decimal commas and thousands separators.

### Maintenance

- Removed duplicated `(1)` workspace copies and stale generated/cache duplicates.

## v0.1.1

Released as the first patch update after `v0.1.0`. This release focuses on ledger correctness and export/import stability.

### Fixed

- Corrected ledger totals to use the stored signed amount as an algebraic sum.
- Improved PDF export font handling for Chinese, Chinese punctuation, and mixed Chinese/Latin text.
- Fixed clipboard import parsing so each strongly detected field is matched only once per row.
- Improved amount display consistency, including German decimal comma formatting.

### Changed

- Document export and table totals now follow the same stored amount semantics.

### Developer Notes

- Improved Windows packaging scripts to reduce PyInstaller cleanup permission issues.
- Build scripts now fail correctly when PyInstaller fails, instead of reporting a successful build.
- Build scripts now remove previous `build` and `dist` outputs before invoking PyInstaller.

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
