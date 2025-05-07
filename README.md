# 🔍 aCRF Builder

**Automate XFDF annotations, PDF bookmarks, and Table of Contents generation for FDA-compliant annotated CRFs.**  
Seamlessly combine SAS, PDFtk, and Python into one efficient toolkit.

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
| `anno_location` | ❌ No        | Macro directory | Path to Excel file and blank CRF PDF.            |
| `page_orient`   | ❌ No        | `Portrait`      | PDF page orientation (`Portrait` / `Landscape`). |
| `font_sz`       | ❌ No        | `10`            | Annotation font size.                            |
| `debug`         | ❌ No        | `N`             | Retain intermediate datasets (`Y` / `N`).        |


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
| `FILE_PATH`   | ❌ No        | Macro directory | Path where input/output PDF files are stored. |
| `INPDF`       | ✅ Yes       | —               | Name of input (annotated) PDF file.           |
| `OUTPDF`      | ❌ No        | `aCRF`          | Output PDF file name.                         |
| `Create_TOC`  | ❌ No        | `N`             | Whether to generate a TOC (`Y` / `N`).        |
| `TOC_SIZE`    | ❌ No        | `12`            | Font size for the TOC.                        |


## Example Usage
%add_bookmarks(INPDF=blankCRF_Positioned, Create_TOC=Y);

## 🐍 Python Script: `create_toc_hypl.py`

This script generates a **Table of Contents (TOC)** for a PDF using its bookmarks and inserts it at the beginning. It also preserves and adjusts original bookmarks with accurate page offsets.

### 🔧 Features

- Extracts existing bookmarks from the input PDF.
- Dynamically calculates number of TOC pages based on font size.
- Supports both Portrait and Landscape PDFs.
- Adds clickable TOC hyperlinks to target sections.
- Retains and adjusts original PDF bookmarks after TOC insertion.

## 📥 Script Arguments

| **Argument**   | **Required** | **Type** | **Description**                                                  |
|----------------|--------------|----------|------------------------------------------------------------------|
| `input.pdf`    | ✅ Yes       | File     | Path to the input PDF file that contains bookmarks.              |
| `output.pdf`   | ✅ Yes       | File     | Path where the output PDF (with TOC and bookmarks) will be saved.|
| `font_size`    | ✅ Yes       | Integer  | Font size for the TOC text (e.g., `10`, `12`, `14`).             |

### 📂 Inputs

- Bookmarked input PDF file.
- Font size for TOC entries (recommended: 10–14).

### 📝 Usage

python create_toc_hypl.py input.pdf output.pdf font_size

## 🔄 End-to-End Workflow

Follow these steps to build your FDA-compliant aCRF:

### 🛠️ Step-by-Step Instructions

1. **Add annotation details** in `ANNOTATION.xlsx`:
   - Fill the **Annotation** sheet with domain metadata and positions.
   - Fill the **Length** sheet with character widths per domain/font.

2. **Run `make_xfdf` macro**:
   - This generates the **unpositioned XFDF** file for annotation.

3. **Import XFDF into PDF Editor**:
   - Open your blank CRF PDF.
   - Import the XFDF and **position the annotations correctly**.

4. **Save the positioned annotated PDF**.

5. **Add bookmark details** in the Excel template (if using Excel-driven bookmarks).

6. **Run `add_bookmarks` macro**:
   - This adds bookmarks to the positioned PDF.

7. **Generate TOC (optional)**:
   - Set `Create_TOC=Y` when running `add_bookmarks`.
   - Ensure Python and `create_toc_hypl.py` are available.
