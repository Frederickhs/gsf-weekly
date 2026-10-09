# Changelog

All notable changes to the **Success Factors Weekly Training Report** application are documented in this file.

The project follows [Semantic Versioning](https://semver.org/):

- **MAJOR** version for major or incompatible changes
- **MINOR** version for new features
- **PATCH** version for corrections and minor improvements

## [Unreleased]

### Planned

- Complete real-user testing with another trainer.
- Record issues, usability feedback, and enhancement requests from testing.
- Promote the application to version `1.0.0` after successful external validation.

## [0.2.0] - 2026-10-09

### Added

- Added a **Search SSO** utility that becomes available after a Weekly Training workbook is loaded.
- Added employee lookup across the entire loaded workbook, independent of the selected Park, Department, or Organization filters.
- Added employee details for Technician, SSO / Employee ID, Park, Department, and Organization / Venue.
- Added an overall employee status using the existing 15-day rules.
- Added a pending-training list showing Item Title, Days Until Due, Status, and Organization / Venue.
- Added color-coded Good, Due Soon, and Overdue results.
- Added duplicate suppression for identical pending-training results.
- Added keyboard support so pressing Enter runs the SSO search.
- Added a Not Found result for SSOs that do not exist in the loaded workbook.

### Changed

- Updated the application minimum window size to `980 x 720`.
- Updated the EXE build workflow to archive the previous distribution folder before rebuilding.

## [0.1.0] - 2026-10-09

### Added

- Created the initial **Success Factors Weekly Training Report** desktop application.
- Added support for selecting a Weekly Training `.xlsx` or `.xlsm` workbook from any accessible folder.
- Added automatic worksheet and header-row detection.
- Added source-column recognition for:
  - Employee ID / SSO
  - Display Name
  - First Name
  - Last Name
  - Item Title
  - Days Until Due
  - Park Description
  - Organization Description
  - Department / Sub Division Description
- Added cascading filters for:
  - Park
  - Department
  - Organization / Venue
- Added searchable Organization / Venue selection.
- Added multi-select support without requiring the Ctrl key.
- Added **Select All Displayed** and **Clear All** controls.
- Added a dedicated **Selected** panel that updates automatically.
- Added unique employee and selected-organization counts.
- Added a visible progress bar during report loading and generation.
- Added Excel report generation with:
  - Summary worksheet
  - One worksheet per selected Organization / Venue
  - Technician first and last name when available
  - SSO / Employee ID
  - Item Title
  - Days Until Due
  - Status
- Added status rules:
  - **Overdue** when Days Until Due is less than `0`
  - **Due Soon** when Days Until Due is from `0` through `15`
  - **Good** when the employee has no overdue or due-soon item
- Added professional workbook formatting, filters, frozen panes, colors, and column widths.
- Added automatic archiving of existing Output contents before generating a new report.
- Added a double-click development launcher.
- Added a PyInstaller one-folder EXE build workflow.
- Added EXE-safe Output and Archive paths so generated files remain beside the distributed application.
- Added automatic archiving of the previous distribution folder before rebuilding the EXE.
- Added `VERSION.txt` to identify the application version.
- Added an end-user `README.txt` to the release folder.
- Added a developer and maintenance guide covering common changes, testing, rebuilding, Git, and troubleshooting.

### Changed

- Standardized the application, EXE, workbook, and release-folder name as **Success Factors Weekly Training Report**.
- Changed technician display values to use First Name plus Last Name when available, with Display Name as fallback.
- Removed the redundant **Load Report** button because Browse automatically loads the selected workbook.
- Removed **Action Items** from the Summary worksheet.
- Updated the application minimum window size to `980 x 720`.
- Updated the build workflow to preserve prior application builds in `Archive\Application Builds`.

### Fixed

- Corrected Organization / Venue selection and unselection behavior.
- Prevented search filtering from clearing existing Organization / Venue selections.
- Corrected automatic refresh behavior in the Selected panel.
- Corrected duplicate or conflicting application names across the interface, workbook, EXE, build files, and distribution folder.
- Corrected frozen-application paths for Output and Archive folders.
- Preserved SSO / Employee ID values as text in generated workbooks.

### Technical

- Python `3.14.8`
- pandas `3.0.6`
- openpyxl `3.1.5`
- PyInstaller `6.22.3`
- Private Git repository: [GSF Weekly](https://github.com/Frederickhs/gsf-weekly.git)

---

## Release checklist

Before adding a new version to this changelog:

1. Test the Python development version.
2. Generate and inspect a workbook.
3. Rebuild the one-folder EXE.
4. Test the rebuilt EXE.
5. Update `VERSION.txt`.
6. Update `README.txt` when user instructions change.
7. Move completed items from **Unreleased** into a new version section.
8. Commit and push the source changes.
9. Create and push a Git tag for stable releases.
