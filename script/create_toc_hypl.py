# ****************************************************************************************************
# Script Name    : create_toc_hypl.py
#
# Purpose        : Adds a clickable Table of Contents (TOC) to an existing PDF using its
#                  existing bookmarks. Uses PyMuPDF to insert TOC pages and hyperlink them to
#                  corresponding document sections.
#
# Author         : Manivannan Mathialagan
# Created On     : 30-Apr-2025
#
# Example Usage
#   python create_toc_hypl.py blankCRF_Positioned.pdf aCRF.pdf 12
#
# Notes
#   - Requires Python 3 and PyMuPDF installed (`pip install pymupdf`).
# ****************************************************************************************************

import sys
import fitz  # PyMuPDF
import math

def extract_toc_entries(doc):
    toc = doc.get_toc(simple=True)
    entries = []
    for entry in toc:
        level, title, page_num = entry[:3]
        entries.append((level, title.strip(), page_num - 1))  # 0-based page numbers
    return entries

def wrap_text(text, font_size, max_width):
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if fitz.get_text_length(test_line, fontsize=font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return lines

def paginate_wrapped_entries(toc_entries, font_size, max_width, lines_per_page):
    paginated_entries = []
    current_page_entries = []
    current_line_count = 0

    for entry in toc_entries:
        level, title, target_page = entry
        indent = 20 * (level - 1)
        available_width = max_width - indent
        wrapped_lines = wrap_text(title, font_size, available_width)
        line_count = len(wrapped_lines)

        if current_line_count + line_count > lines_per_page:
            paginated_entries.append(current_page_entries)
            current_page_entries = []
            current_line_count = 0

        current_page_entries.append((entry, wrapped_lines))
        current_line_count += line_count

    if current_page_entries:
        paginated_entries.append(current_page_entries)

    return paginated_entries

def generate_toc_pages(paginated_entries, font_size, page_width, page_height):
    toc_doc = fitz.open()
    link_targets = []
    left_margin = 50
    right_margin = 60
    top_margin = 50
    y_spacing = font_size * 1.5

    toc_page_count = len(paginated_entries)

    for page_index, entries in enumerate(paginated_entries):
        page = toc_doc.new_page(width=page_width, height=page_height)
        y = top_margin
        for (level, title, target_page), wrapped_lines in entries:
            indent = 20 * (level - 1)
            x = left_margin + indent
            page_number_str = str(target_page + toc_page_count + 1)
            page_number_width = fitz.get_text_length(page_number_str, fontsize=font_size)
            max_x_for_dots = page_width - right_margin - page_number_width - 5

            first_line_y = y  # Needed for hyperlink rectangle

            for i, line in enumerate(wrapped_lines):
                line_width = fitz.get_text_length(line, fontsize=font_size)
                dots = ''
                if i == len(wrapped_lines) - 1:
                    dots_space = max_x_for_dots - (x + line_width + 10)
                    dot_count = max(0, int(dots_space / fitz.get_text_length('.', fontsize=font_size)))
                    dots = '.' * dot_count

                    # Draw line with dots and page number
                    page.insert_text((x, y), f"{line} {dots}", fontsize=font_size)
                    page.insert_text((page_width - right_margin - page_number_width, y), page_number_str, fontsize=font_size)
                else:
                    # Draw line without dots/page number
                    page.insert_text((x, y), line, fontsize=font_size)

                y += y_spacing

            rect = fitz.Rect(x, first_line_y - font_size, page_width - right_margin, y)
            link_targets.append((page_index, rect, target_page))

    return toc_doc, link_targets

def add_toc_hyperlinks(doc, link_targets, toc_page_count):
    for toc_page_index, rect, target_page in link_targets:
        doc[toc_page_index].insert_link({
            "kind": fitz.LINK_GOTO,
            "from": rect,
            "page": target_page + toc_page_count
        })

def shift_bookmark_pages(bookmarks, offset):
    shifted = []
    for bm in bookmarks:
        if len(bm) >= 3:
            level, title, page_num = bm[:3]
            shifted.append([level, title, page_num + offset])
    return shifted

def add_existing_bookmarks(doc, bookmarks, offset):
    shifted = shift_bookmark_pages(bookmarks, offset)
    doc.set_toc(shifted)

def main():
    if len(sys.argv) < 4:
        print("Usage: python create_toc_hypl.py input.pdf output.pdf font_size")
        return

    input_pdf = sys.argv[1]
    output_pdf = sys.argv[2]
    font_size = int(sys.argv[3])
    lines_per_page = 38

    print("Opening original PDF...")
    original = fitz.open(input_pdf)

    print("Extracting original TOC...")
    toc_entries = extract_toc_entries(original)
    original_bookmarks = original.get_toc()

    width, height = fitz.paper_size("a4")
    max_width = width - 120  # account for left/right margins

    print("Paginating wrapped TOC entries...")
    paginated_entries = paginate_wrapped_entries(toc_entries, font_size, max_width, lines_per_page)
    toc_page_count = len(paginated_entries)

    print(f"Generating TOC pages with hyperlinks (Pages: {toc_page_count})...")
    toc_pdf, link_targets = generate_toc_pages(paginated_entries, font_size, width, height)

    print("Merging TOC and original PDF...")
    final = fitz.open()
    final.insert_pdf(toc_pdf)
    final.insert_pdf(original)

    print("Adding TOC hyperlinks...")
    add_toc_hyperlinks(final, link_targets, toc_page_count)

    print("Preserving original bookmarks...")
    add_existing_bookmarks(final, original_bookmarks, toc_page_count)

    print("Saving final PDF...")
    final.save(output_pdf)
    final.close()
    print(f"Done. Output saved to {output_pdf}")

if __name__ == "__main__":
    main()
