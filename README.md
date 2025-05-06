# SAS PDF Annotation Toolkit

A set of SAS macros to automate the creation of FDA-compliant annotated CRFs (aCRFs), including XFDF annotations, PDF bookmarks, and an optional Table of Contents (TOC) using PDFtk and Python.

## 📦 Included Macros

### 1. make_xfdf

Creates XFDF annotation files from an Excel-based input to overlay field-level metadata on blank CRFs.

**Features:**
- Supports customization of font size, font color, border color, and font type per domain.
- Handles multi-page annotation and page orientation.
- Integrates annotation character length logic to position text neatly within fields.

**Input Requirements:**
- ANNOTATION.xlsx with:
    - Annotation sheet: Contains metadata field positions and values.
    - Length sheet: Character width reference per domain or font type.
- Blank CRF PDF file.

## Macro Parameters:**

| **Parameter**   | **Required** | **Default**     | **Description**                                  |
| --------------- | ------------ | --------------- | ------------------------------------------------ |
| `anno_location` | No           | Macro directory | Path to Excel file and blank CRF PDF.            |
| `page_orient`   | No           | `Portrait`      | PDF page orientation (`Portrait` / `Landscape`). |
| `font_sz`       | No           | `10`            | Annotation font size.                            |
| `debug`         | No           | `N`             | Retain intermediate datasets (`Y` / `N`).        |


## Example Usage

%make_xfdf(font_sz=9, debug=Y);

### 2. add_bookmarks

Adds bookmarks to a PDF and optionally generates a front-page Table of Contents (TOC) using Python and PDFtk.

**Features:**
- Bookmark navigation structure for easier review of aCRF.
- TOC generation with customizable font size.
- Automatically merges TOC with annotated CRF.

## Requirements

- **SAS** (tested with SAS 9.4)
- **PDFtk** (command-line version)
- **Python 3+** (optional, for TOC generation)
- OS: Windows or Mac (macros use system calls)

Ensure `PDFtk` and `Python` are available in your system.

## Macro Parameters
| **Parameter** | **Required** | **Default**     | **Description**                               |
| ------------- | ------------ | --------------- | --------------------------------------------- |
| `FILE_PATH`   | No           | Macro directory | Path where input/output PDF files are stored. |
| `INPDF`       | Yes          | —               | Name of input (annotated) PDF file.           |
| `OUTPDF`      | No           | `aCRF`          | Output PDF file name.                         |
| `Create_TOC`  | No           | `N`             | Whether to generate a TOC (`Y` / `N`).        |
| `TOC_SIZE`    | No           | `12`            | Font size for the TOC.                        |


## Example Usage
%add_bookmarks(INPDF=blankCRF_Positioned, Create_TOC=Y);

