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
    return [(level, title.strip(), page_num - 1) for level, title, page_num in toc]

def wrap_text(text, fontsize, max_width):
    words = text.split(' ')
    lines = []
    current_line = words[0]
    for word in words[1:]:
        if fitz.get_text_length(current_line + ' ' + word, fontsize=fontsize) <= max_width:
            current_line += ' ' + word
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines

def paginate_wrapped_entries(toc_entries, font_size, page_width, page_height, left_margin, right_margin, top_margin, bottom_margin):
    """Split entries across pages based on wrapped line count"""
    usable_height = page_height - top_margin - bottom_margin
    line_height = font_size * 1.5
    lines_per_page = int(usable_height // line_height)

    max_title_width = page_width - left_margin - right_margin - 80  # Adjusted to wrap earlier
    pages = []
    current_page = []
    used_lines = 0

    for entry in toc_entries:
        level, title, page_num = entry
        indent = 20 * (level - 1)
        wrapped_lines = wrap_text(title, font_size, max_title_width - indent)
        line_count = len(wrapped_lines)

        if used_lines + line_count > lines_per_page:
            pages.append(current_page)
            current_page = []
            used_lines = 0

        current_page.append((entry, wrapped_lines))
        used_lines += line_count

    if current_page:
        pages.append(current_page)

    return pages

def generate_toc_pages(pages, font_size, page_size):
    toc_doc = fitz.open()
    link_targets = []
    page_width, page_height = page_size
    left_margin = 50
    right_margin = 60
    top_margin = 50
    y_spacing = font_size * 1.5

    for page_index, entries in enumerate(pages):
        page = toc_doc.new_page(width=page_width, height=page_height)
        y = top_margin
        for (level, title, target_page), wrapped_lines in entries:
            indent = 20 * (level - 1)
            x = left_margin + indent
            page_number_str = str(target_page + len(pages) + 1)
            page_number_width = fitz.get_text_length(page_number_str, fontsize=font_size)
            max_x_for_dots = page_width - right_margin - page_number_width - 5

            for i, line in enumerate(wrapped_lines):
                line_width = fitz.get_text_length(line, fontsize=font_size)
                if i == len(wrapped_lines) - 1:
                    dots_space = max_x_for_dots - (x + line_width + 10)
                    dot_count = max(0, int(dots_space / fitz.get_text_length('.', fontsize=font_size)))
                    dots = '.' * dot_count
                else:
                    dots = ''
                page.insert_text((x, y), f"{line} {dots}", fontsize=font_size)
                if i == len(wrapped_lines) - 1:
                    page.insert_text((page_width - right_margin - page_number_width, y), page_number_str, fontsize=font_size)
                y += y_spacing

            rect = fitz.Rect(x, y - len(wrapped_lines)*y_spacing, page_width - right_margin, y)
            link_targets.append((page_index, rect, target_page))

    return toc_doc, link_targets

def add_toc_hyperlinks(doc, link_targets, toc_pages):
    for toc_page_index, rect, target_page in link_targets:
        doc[toc_page_index].insert_link({
            "kind": fitz.LINK_GOTO,
            "from": rect,
            "page": target_page + toc_pages
        })

def shift_bookmark_pages(bookmarks, offset):
    return [[lvl, title, pg + offset] for lvl, title, pg in bookmarks if len([lvl, title, pg]) >= 3]

def add_existing_bookmarks(doc, bookmarks, offset):
    doc.set_toc(shift_bookmark_pages(bookmarks, offset))

def main():
    if len(sys.argv) < 4:
        print("Usage: python script.py input.pdf output.pdf font_size")
        return

    input_pdf = sys.argv[1]
    output_pdf = sys.argv[2]
    font_size = int(sys.argv[3])

    print("Opening original PDF...")
    original = fitz.open(input_pdf)
    original_toc = extract_toc_entries(original)
    original_bookmarks = original.get_toc()

    page_width, page_height = original[0].rect.width, original[0].rect.height
    print(f"Page size: {page_width} x {page_height}")

    print("Paginating TOC based on wrapped lines...")
    left_margin = 50
    right_margin = 60
    top_margin = 50
    bottom_margin = 50
    pages = paginate_wrapped_entries(
        original_toc,
        font_size,
        page_width,
        page_height,
        left_margin,
        right_margin,
        top_margin,
        bottom_margin
    )
    toc_pages = len(pages)
    print(f"TOC will span {toc_pages} page(s)")

    print("Generating TOC pages...")
    toc_pdf, link_targets = generate_toc_pages(pages, font_size, (page_width, page_height))

    print("Merging TOC and original PDF...")
    final = fitz.open()
    final.insert_pdf(toc_pdf)
    final.insert_pdf(original)

    print("Adding TOC hyperlinks...")
    add_toc_hyperlinks(final, link_targets, toc_pages)

    print("Preserving original bookmarks...")
    add_existing_bookmarks(final, original_bookmarks, toc_pages)

    print("Saving final PDF...")
    final.save(output_pdf)
    final.close()
    print(f"Done. Output saved to {output_pdf}")

if __name__ == "__main__":
    main()
