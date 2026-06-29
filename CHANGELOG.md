# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and the versioning follows [Semantic Versioning](https://semver.org/).

---
## [v2.0.1] - 2026-06-29

### Changed
- Version number to test the whole build process.

## [v2.0.0] - 2026-06-29

### Added
- Rebuilt the desktop UI with PySide6/Qt.
- Added portable Python release packaging for managed Windows devices.
- Added GitHub Actions release automation for version tags.
- Added nested array grouping, unfolding, parent metadata, table preview, and transform options.
- Added app/tray icon assets.

### Changed
- Reworked export internals to avoid a pandas runtime dependency.
- Updated build outputs so generated release zips are ignored by git and attached through GitHub Releases.

## [v1.0.10] - 2025-08-11

### Added
- Support for selecting and exporting deeply nested arrays using index-based path notation (e.g., `orders[0].items`).
- Recursive array discovery: all arrays, including those nested inside objects or arrays, are now available for export.
- Updated README to explain the array path notation and how to select nested arrays.

## [v1.0.9] - 2025-08-08
### Added
- Initial public release
- Standalone `.exe` build with GUI for JSON-to-Excel conversion
- Sample data and README included in release
