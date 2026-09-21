# BioReviewPy

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22865427.svg)](https://doi.org/10.5281/zenodo.22865427)

**BioReviewPy** is an open-source Python application designed to support bibliographic data management and screening workflows in biomedical systematic reviews and evidence-synthesis projects.

The software integrates bibliographic file importation, metadata harmonization, duplicate detection, manual verification of uncertain matches, record classification, screening support, audit tracking, and export of cleaned datasets within a single workflow. Its deduplication strategy is intentionally conservative: uncertain matches are retained for researcher review rather than being automatically deleted.

## Main features

- Automatic detection of supported bibliographic databases and compatible export formats
- Importation of records from multiple bibliographic sources
- Metadata cleaning and harmonization
- Automatic detection of exact duplicates
- Fuzzy matching for potentially duplicated records
- Manual confirmation or rejection of uncertain duplicate matches
- Identification and classification of systematic reviews and meta-analyses
- Sequential title/abstract screening support
- PRISMA-oriented record tracking
- Reproducibility and audit reports
- Export of cleaned datasets and audit tables
- Bibliometric dataset preparation for selected workflows
- Multilingual interface: English, Portuguese, and Spanish

## Supported databases and recognized import formats

| Database | Recognized/importable formats |
|---|---|
| PubMed | NBIB, TXT, CSV |
| Web of Science | TXT/CIW, RIS, CSV, XLSX, XLS, HTML, BIB |
| Scopus | TXT, RIS, CSV, BIB |
| Embase | RIS, TXT, CSV, XML, DOCX, XLSX, PDF |
| LILACS/BVS | RIS, TXT, CSV |
| Cochrane Library | TXT |
| SciELO | BIB |

BioReviewPy attempts to identify the source database and compatible file structure automatically, reducing the need for manual file conversion before processing.

## General workflow

1. Import bibliographic files
2. Standardize and harmonize metadata
3. Detect exact and potential duplicates
4. Manually verify uncertain duplicate matches
5. Generate the deduplicated dataset
6. Classify and screen records when required
7. Export cleaned datasets, PRISMA-oriented outputs, and audit reports

## Current version

Current development version: **0.23**

BioReviewPy is under active development. The version used in scientific validation and the corresponding archived release should be cited to ensure reproducibility.

## Installation

BioReviewPy can be run directly from the Python source file.

```bash
pip install -r requirements.txt
python BioReviewPy.py
```

Some import formats use optional local components or libraries. In particular, PDF extraction requires `pypdf` or PyMuPDF, and legacy `.xls` files may require `xlrd` or a local LibreOffice installation.

## Citation

If you use BioReviewPy in research, please cite the archived software release:

**Martins dos Santos, R. (2026). BioReviewPy (Version 0.23) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.22865556**

DOI: **10.5281/zenodo.22865556**

Citation metadata are also provided in [`CITATION.cff`](CITATION.cff).

## License

BioReviewPy is distributed under the **MIT License**. See [`LICENSE`](LICENSE) for details.

Copyright © 2026 Rodrigo Martins dos Santos.

## Disclaimer

BioReviewPy is intended to support systematic review and evidence-synthesis workflows. Automated classifications and duplicate-detection results should be verified by researchers whenever appropriate. Users remain responsible for methodological decisions, study eligibility assessment, and final data interpretation.

## Developer and maintainer

**Rodrigo Martins dos Santos**  
E-mail: rodrigoms13@hotmail.com  
Repository: https://github.com/rodrigomartins1391-blip/BioReviewPy

## Version history

Major changes are documented in [`CHANGELOG.md`](CHANGELOG.md).
