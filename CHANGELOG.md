# Changelog

All notable changes to BioReviewPy will be documented in this file.

## [0.24.4] - 2026-09-27

### Added
- Dedicated title/abstract study-screening workflow with `INCLUDE`, `EXCLUDE`, and `MAYBE` decisions.
- Standardized exclusion-reason selector with Portuguese, English, and Spanish labels, plus a free-text “Other reason” option.
- PRISMA 2020 preview window integrated into the study-selection workflow.
- High-resolution PRISMA export in PNG (300 DPI), JPG, and PDF formats.
- Separate export of PRISMA flow data for auditability.

### Changed
- Reorganized the Tools menu to separate the main study-selection workflow from auxiliary bibliographic classification tools.
- Systematic-review/meta-analysis identification is now explicitly treated as an auxiliary classification step rather than an exclusion step.
- Redesigned the PRISMA figure with larger stage labels, improved typography, simplified boxes, and a cleaner publication-oriented layout.
- Simplified the PRISMA diagram to report duplicate removal and study-selection counts without automation-specific removal boxes or full-text exclusion reasons in the figure.
- Updated software version metadata to 0.24.4.

### Notes
- Screening decisions and exclusion reasons remain researcher-controlled and are stored in the project/audit outputs.
- The PRISMA visualization is intended to support reporting and does not replace methodological assessment by reviewers.

## [0.24.0] - 2026-09-26

### Added
- Native support for ERIC (Education Resources Information Center) exports in `.nbib` format.
- Automatic distinction between ERIC NBIB files and PubMed/MEDLINE NBIB files.
- Dedicated parsing of ERIC identifiers (`EJ` and `ED`), titles, authors, abstracts, journals, publication years, and DOIs.
- Support for loading multiple ERIC export batches and combining them into a single ERIC dataset.
- Protection against duplicate ERIC records when export batches overlap.

### Changed
- Updated database identification and export handling to include ERIC.
- Updated software version metadata to 0.24.0.

### Notes
- This release is especially useful for ERIC searches exported in multiple batches.

## [0.23.1] - 2026-09-20

### Fixed
- Added the main `BioReviewPy.py` source file to the public repository and release workflow.
- Updated software version metadata for consistency across the source code and citation files.

### Notes
- This is a corrective release of v0.23. No changes were made to the core deduplication, screening, or bibliographic processing algorithms.

## [0.23] - 2026-09-20

### Added
- Multilingual interface in Portuguese, English, and Spanish.
- Automatic database and export-format detection for supported sources.
- Semi-automated and auditable duplicate-detection workflow.
- Manual review of uncertain duplicate matches.
- Project saving and reopening.
- PRISMA-oriented record tracking and export.
- Reproducibility/audit reports.
- Systematic review and meta-analysis classification support.
- Bibliometric dataset preparation tools.
- Software authorship, license, repository, and citation information in the interface and exported reports.

### Notes
- This release remains under active development and should be validated with additional datasets and export variants.
