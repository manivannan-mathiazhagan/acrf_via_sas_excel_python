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

def generate_toc_pages(toc_entries, font_size, lines_per_page, toc_page_count):
    toc_doc = fitz.open()
    link_targets = []
    width, height = fitz.paper_size("a4")
    total_entries = len(toc_entries)
    total_pages = math.ceil(total_entries / lines_per_page)

    for i in range(total_pages):
        page = toc_doc.new_page(width=width, height=height)
        start = i * lines_per_page
        end = min(start + lines_per_page, total_entries)

        for j, (level, title, target_page) in enumerate(toc_entries[start:end]):
            y = 50 + j * font_size * 1.5
            displayed_page = target_page + toc_page_count + 1
            indent = 20 * (level - 1)
            title_x = 50 + indent
            max_page_number_x = width - 60  # Right margin for page numbers

            # Measure text widths
            text_width = fitz.get_text_length(title, fontsize=font_size)
            page_num_text = str(displayed_page)
            page_num_width = fitz.get_text_length(page_num_text, fontsize=font_size)

            # Determine dot space
            dots_space = max_page_number_x - page_num_width - (title_x + text_width + 10)
            num_dots = max(2, int(dots_space / fitz.get_text_length(".", fontsize=font_size)) - 6)
            dots = "." * num_dots

            # Insert title + dots
            page.insert_text((title_x, y), f"{title} {dots}", fontsize=font_size)

            # Right-align page number
            page.insert_text((max_page_number_x - page_num_width, y), page_num_text, fontsize=font_size)

            # Create link rectangle
            rect = fitz.Rect(title_x, y - font_size, max_page_number_x, y + 5)
            link_targets.append((i, rect, target_page))

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
    page_height = max(width, height)  # always treat larger side as height

    top_margin = 50
    bottom_margin = 50
    usable_height = page_height - top_margin - bottom_margin
    line_height = font_size * 1.5
    lines_per_page = int(usable_height // line_height)

    toc_page_count = math.ceil(len(toc_entries) / lines_per_page)

    print(f"Page orientation: {'Portrait' if portrait else 'Landscape'}")
    print(f"Calculated lines per page: {lines_per_page}")
    print(f"TOC will span {toc_page_count} page(s)")

    print("Generating TOC pages with hyperlinks...")
    toc_pdf, link_targets = generate_toc_pages(toc_entries, font_size, lines_per_page, toc_page_count)

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
