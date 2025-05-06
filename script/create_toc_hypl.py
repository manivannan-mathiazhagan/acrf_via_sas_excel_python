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

def wrap_text(text, max_width, font_size):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = current_line + " " + word if current_line else word
        if fitz.get_text_length(test_line, fontsize=font_size) < max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)
    return lines

def generate_toc_pages(toc_entries, font_size, lines_per_page, toc_page_count):
    toc_doc = fitz.open()
    link_targets = []
    width, height = fitz.paper_size("a4")
    line_height = font_size * 1.5

    i = 0
    page = toc_doc.new_page(width=width, height=height)
    current_line = 0
    total_lines = lines_per_page

    for entry in toc_entries:
        level, title, target_page = entry
        indent = 20 * (level - 1)
        title_x = 50 + indent
        max_page_number_x = width - 60
        max_text_width = max_page_number_x - title_x - 60  # Leave room for page numbers and dots

        wrapped_lines = wrap_text(title, max_text_width, font_size)
        num_lines = len(wrapped_lines)

        # Create new page if needed
        if current_line + num_lines > total_lines:
            page = toc_doc.new_page(width=width, height=height)
            current_line = 0
            i += 1

        top_y = 50 + current_line * line_height
        bottom_y = top_y + num_lines * line_height
        rect = fitz.Rect(title_x, top_y - font_size, max_page_number_x, bottom_y)
        link_targets.append((i, rect, target_page))

        for k, line in enumerate(wrapped_lines):
            y = 50 + (current_line * line_height)
            displayed_page = target_page + toc_page_count + 1
            text_width = fitz.get_text_length(line, fontsize=font_size)
            page_num_text = str(displayed_page)
            page_num_width = fitz.get_text_length(page_num_text, fontsize=font_size)

            # Dots only on last line
            if k == num_lines - 1:
                dots_space = max_page_number_x - page_num_width - (title_x + text_width + 10)
                num_dots = max(2, int(dots_space / fitz.get_text_length(".", fontsize=font_size)) - 6)
                dots = "." * num_dots
                page.insert_text((title_x, y), f"{line} {dots}", fontsize=font_size)
                page.insert_text((max_page_number_x - page_num_width, y), page_num_text, fontsize=font_size)
            else:
                page.insert_text((title_x, y), line, fontsize=font_size)

            current_line += 1

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

    print("Opening original PDF...")
    original = fitz.open(input_pdf)

    print("Extracting original TOC...")
    toc_entries = extract_toc_entries(original)
    original_bookmarks = original.get_toc()

    # Dynamically calculate lines per page
    first_page = original[0]
    width, height = first_page.rect.width, first_page.rect.height
    portrait = height > width
    page_height = max(width, height)  # treat larger side as height

    top_margin = 50
    bottom_margin = 50
    usable_height = page_height - top_margin - bottom_margin
    line_height = font_size * 1.5
    lines_per_page = int(usable_height // line_height)

    toc_page_count = 0  # Temp, will recalculate after wrapping
    print(f"Page orientation: {'Portrait' if portrait else 'Landscape'}")
    print(f"Initial estimate - lines per page: {lines_per_page}")

    print("Generating TOC pages with hyperlinks...")
    toc_pdf, link_targets = generate_toc_pages(toc_entries, font_size, lines_per_page, toc_page_count)
    toc_page_count = len(toc_pdf)  # Accurate count now

    print(f"TOC will span {toc_page_count} page(s)")

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
