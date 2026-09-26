# Changelog

All notable changes to BioReviewPy will be documented in this file.

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
