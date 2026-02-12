"""
Convert Ocean_From_Motion.docx into a static HTML book in docs/.
"""

import os
import re
import shutil
from html import escape
from docx import Document
from docx.oxml.ns import qn

DOCX_PATH = os.path.join(os.path.dirname(__file__),
                         "Ocean_From_Motion (1).docx")
DOCS_DIR = os.path.join(os.path.dirname(__file__), "docs")
BASE_URL = "/Ocean-From-Motion"


# ── Helpers ──────────────────────────────────────────────────────────

def runs_to_html(paragraph):
    """Convert a paragraph's runs into an HTML string, preserving bold/italic."""
    parts = []
    for run in paragraph.runs:
        text = escape(run.text)
        if not text:
            continue
        if run.bold and run.italic:
            text = f"<strong><em>{text}</em></strong>"
        elif run.bold:
            text = f"<strong>{text}</strong>"
        elif run.italic:
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts) or escape(paragraph.text)


def table_to_html(table):
    """Convert a docx table to an HTML <table>."""
    rows_html = []
    for ri, row in enumerate(table.rows):
        cells = []
        tag = "th" if ri == 0 else "td"
        for cell in row.cells:
            cells.append(f"<{tag}>{escape(cell.text)}</{tag}>")
        rows_html.append("<tr>" + "".join(cells) + "</tr>")
    return "<table>\n" + "\n".join(rows_html) + "\n</table>"


def is_section_heading(paragraph):
    """Detect numbered sub-section headings like '4.1 Title' in Body Text."""
    text = paragraph.text.strip()
    if not text:
        return False
    # Patterns: "1. Title", "4.1 Title", "4.1.1 Title"
    if re.match(r"^\d+(\.\d+)*\.?\s", text):
        # Check if whole paragraph is bold
        if paragraph.runs and all(r.bold for r in paragraph.runs if r.text.strip()):
            return True
        # Or if style is Body Text and it looks like a heading
        if paragraph.style.name == "Body Text":
            return True
    return False


def classify_heading_level(text):
    """Determine h2/h3/h4 level from numbered heading text."""
    m = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?\.", text.strip())
    if not m:
        m = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?\s", text.strip())
    if m:
        if m.group(3):
            return "h4"
        elif m.group(2):
            return "h3"
    return "h2"


# ── Document parsing ────────────────────────────────────────────────

def parse_document(doc):
    """
    Walk the document body and produce an ordered list of elements:
    each is ('paragraph', para) or ('table', table).
    """
    elements = []
    para_iter = iter(doc.paragraphs)
    table_iter = iter(doc.tables)

    # Map XML elements to python-docx objects
    para_map = {}
    table_map = {}
    for p in doc.paragraphs:
        para_map[id(p._element)] = p
    for t in doc.tables:
        table_map[id(t._element)] = t

    body = doc.element.body
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            p = para_map.get(id(child))
            if p:
                elements.append(("paragraph", p))
        elif tag == "tbl":
            t = table_map.get(id(child))
            if t:
                elements.append(("table", t))
    return elements


def split_into_sections(elements):
    """
    Split elements into logical sections:
    - front_matter (indices 0-76)
    - preface (indices 77-91, the actual preface text)
    - chapters (Chapter 0-14, splitting on Heading 1)
    - appendix_a, appendix_b
    """
    sections = {}

    # First, find all paragraph indices for key boundaries
    para_idx = 0
    elem_indices = {}  # para_index -> element_index
    for ei, (etype, obj) in enumerate(elements):
        if etype == "paragraph":
            elem_indices[para_idx] = ei
            para_idx += 1

    # Find key paragraph indices
    heading1_indices = []  # (para_idx, elem_idx, text)
    para_idx = 0
    for ei, (etype, obj) in enumerate(elements):
        if etype == "paragraph":
            if obj.style.name.startswith("Heading"):
                heading1_indices.append((para_idx, ei, obj.text))
            para_idx += 1

    # Preface: paragraphs 77 through 91 (actual preface content, skip embedded TOC)
    preface_start = elem_indices.get(77, 0)
    preface_end = elem_indices.get(92, 0)  # exclusive
    sections["preface"] = elements[preface_start:preface_end]

    # Find chapter boundaries from Heading 1 paragraphs
    # Part headers are immediately followed by chapter headings, so we skip them
    chapter_starts = []  # (elem_idx, chapter_id, title)
    part_names = {}  # chapter_num -> part_name

    current_part = None
    for para_idx, ei, text in heading1_indices:
        text_stripped = text.strip()
        if text_stripped.startswith("PART"):
            current_part = text_stripped
            continue
        m = re.match(r"Chapter\s+(\d+)", text_stripped)
        if m:
            ch_num = int(m.group(1))
            chapter_starts.append((ei, ch_num))
            if current_part:
                part_names[ch_num] = current_part

    # Build chapter sections
    for idx, (ei, ch_num) in enumerate(chapter_starts):
        if idx + 1 < len(chapter_starts):
            next_ei = chapter_starts[idx + 1][0]
            # If the next chapter's heading is preceded by a PART heading,
            # we need to end this chapter before the PART heading
            # Find the PART heading element index
            end_ei = next_ei
            # Check if there's a PART heading right before next chapter
            for pi, pei, ptext in heading1_indices:
                if pei == next_ei - 1 and ptext.strip().startswith("PART"):
                    end_ei = pei
                    break
                # Sometimes PART heading IS at next_ei - look back
            sections[f"chapter-{ch_num}"] = elements[ei:end_ei]
        else:
            # Last chapter: goes until appendix A
            # Find "APPENDIX A" paragraph
            app_a_ei = None
            pidx = 0
            for eei, (etype, obj) in enumerate(elements):
                if etype == "paragraph":
                    if obj.text.strip() == "APPENDIX A":
                        app_a_ei = eei
                        break
                    pidx += 1
            if app_a_ei:
                sections[f"chapter-{ch_num}"] = elements[ei:app_a_ei]
            else:
                sections[f"chapter-{ch_num}"] = elements[ei:]

    # Appendices
    app_a_ei = None
    app_b_ei = None
    pidx = 0
    for eei, (etype, obj) in enumerate(elements):
        if etype == "paragraph":
            if obj.text.strip() == "APPENDIX A":
                app_a_ei = eei
            elif obj.text.strip() == "APPENDIX B":
                app_b_ei = eei

    if app_a_ei and app_b_ei:
        sections["appendix-a"] = elements[app_a_ei:app_b_ei]
    elif app_a_ei:
        sections["appendix-a"] = elements[app_a_ei:]

    if app_b_ei:
        sections["appendix-b"] = elements[app_b_ei:]

    return sections, part_names


def section_to_html(section_elements, skip_count=0):
    """Convert a list of (type, obj) elements into HTML body content.

    skip_count: number of leading non-empty paragraphs to skip
    (used to skip heading + subtitle that are rendered separately).
    """
    html_parts = []
    skipped = 0

    for etype, obj in section_elements:
        if etype == "table":
            html_parts.append(table_to_html(obj))
            continue

        # It's a paragraph
        p = obj
        text = p.text.strip()
        if not text:
            continue

        # Skip the first N non-empty paragraphs
        if skipped < skip_count:
            skipped += 1
            continue

        # Heading 1 style
        if p.style.name.startswith("Heading"):
            html_parts.append(f"<h1>{runs_to_html(p)}</h1>")
            continue

        # Numbered sub-section headings
        if is_section_heading(p):
            level = classify_heading_level(text)
            html_parts.append(f"<{level}>{runs_to_html(p)}</{level}>")
            continue

        # Regular paragraphs (both Normal and Body Text)
        content = runs_to_html(p)

        # Detect bullet-like paragraphs (starting with bullet character or •)
        if text.startswith("\u2022") or text.startswith("- "):
            html_parts.append(f"<p class=\"list-item\">{content}</p>")
        else:
            html_parts.append(f"<p>{content}</p>")

    return "\n".join(html_parts)


# ── Chapter metadata ────────────────────────────────────────────────

CHAPTER_TITLES = {
    0: "A Universe of Motion",
    1: "Heat",
    2: "Polarity",
    3: "Existence",
    4: "Righteousness",
    5: "Order",
    6: "Movement",
    7: "Entropy",
    8: "Learning Systems",
    9: "Identity and Persistence",
    10: "Agency and Choice",
    11: "Error, Correction, and Growth",
    12: "Ethics as Stability",
    13: "Coercion as Forced Motion",
    14: "Freedom as Available Motion",
}

PARTS = {
    0: ("I", "Orientation"),
    1: ("II", "The Motion Calendar"),
    2: ("II", "The Motion Calendar"),
    3: ("II", "The Motion Calendar"),
    4: ("II", "The Motion Calendar"),
    5: ("II", "The Motion Calendar"),
    6: ("II", "The Motion Calendar"),
    7: ("III", "Systems"),
    8: ("III", "Systems"),
    9: ("III", "Systems"),
    10: ("III", "Systems"),
    11: ("III", "Systems"),
    12: ("IV", "Meaning"),
    13: ("IV", "Meaning"),
    14: ("IV", "Meaning"),
}

# Navigation order
NAV_ORDER = (
    ["preface"]
    + [f"chapter-{i}" for i in range(15)]
    + ["appendix-a", "appendix-b"]
)

NAV_LABELS = {"preface": "Preface", "appendix-a": "Appendix A", "appendix-b": "Appendix B"}
for i in range(15):
    NAV_LABELS[f"chapter-{i}"] = f"Chapter {i}"


# ── HTML Templates ──────────────────────────────────────────────────

def page_html(title, body_content, page_id, subtitle=None):
    """Wrap body content in a full HTML page with nav."""
    idx = NAV_ORDER.index(page_id) if page_id in NAV_ORDER else -1
    prev_link = ""
    next_link = ""
    if idx > 0:
        prev_id = NAV_ORDER[idx - 1]
        prev_link = f'<a href="{BASE_URL}/{prev_id}.html">&larr; {NAV_LABELS[prev_id]}</a>'
    if 0 <= idx < len(NAV_ORDER) - 1:
        next_id = NAV_ORDER[idx + 1]
        next_link = f'<a href="{BASE_URL}/{next_id}.html">{NAV_LABELS[next_id]} &rarr;</a>'

    subtitle_html = f'<p class="subtitle">{escape(subtitle)}</p>' if subtitle else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{escape(title)} — Ocean From Motion</title>
<link rel="stylesheet" href="{BASE_URL}/style.css">
</head>
<body>
<header>
  <nav class="top-nav">
    <a href="{BASE_URL}/" class="nav-title">Ocean From Motion</a>
    <div class="nav-links">
      <a href="{BASE_URL}/">Contents</a>
      <a href="{BASE_URL}/Ocean_From_Motion.docx" class="download-link">Download .docx</a>
    </div>
  </nav>
</header>
<main>
  <h1 class="page-title">{escape(title)}</h1>
  {subtitle_html}
  {body_content}
</main>
<footer>
  <nav class="chapter-nav">
    <div class="nav-prev">{prev_link}</div>
    <div class="nav-toc"><a href="{BASE_URL}/">Table of Contents</a></div>
    <div class="nav-next">{next_link}</div>
  </nav>
</footer>
</body>
</html>
"""


def index_html():
    """Generate the cover/index page."""
    toc_items = []

    current_part = None
    for ch_num in range(15):
        part_num, part_name = PARTS[ch_num]
        part_key = f"PART {part_num}: {part_name.upper()}"
        if part_key != current_part:
            current_part = part_key
            toc_items.append(f'<li class="part-header">Part {part_num}: {part_name}</li>')
        toc_items.append(
            f'<li><a href="{BASE_URL}/chapter-{ch_num}.html">'
            f"Chapter {ch_num}: {CHAPTER_TITLES[ch_num]}</a></li>"
        )

    toc_items.append(f'<li class="part-header">Appendices</li>')
    toc_items.append(f'<li><a href="{BASE_URL}/appendix-a.html">Appendix A: Mathematical Constants</a></li>')
    toc_items.append(f'<li><a href="{BASE_URL}/appendix-b.html">Appendix B: Notation Reference</a></li>')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ocean From Motion</title>
<link rel="stylesheet" href="{BASE_URL}/style.css">
</head>
<body>
<main class="cover">
  <h1 class="book-title">Ocean From Motion</h1>
  <p class="book-subtitle">A Study in the Nature of Possible Primitives of Motion</p>
  <p class="book-subtitle-secondary">Incorporating<br>The Motion Calendar<br>
  <em>A Universe of Motion Rather Than in Motion</em></p>

  <div class="download-section">
    <a href="{BASE_URL}/Ocean_From_Motion.docx" class="btn-download">Download Original (.docx)</a>
  </div>

  <nav class="toc">
    <h2>Contents</h2>
    <ol>
      <li><a href="{BASE_URL}/preface.html">Preface</a></li>
      {"".join(toc_items)}
    </ol>
  </nav>
</main>
</body>
</html>
"""


# ── Main ─────────────────────────────────────────────────────────────

def main():
    os.makedirs(DOCS_DIR, exist_ok=True)

    print("Reading .docx ...")
    doc = Document(DOCX_PATH)
    elements = parse_document(doc)
    sections, part_names = split_into_sections(elements)

    # Write index
    print("Writing index.html")
    with open(os.path.join(DOCS_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html())

    # Write preface
    # Skip 1 paragraph: the "PREFACE" title (we render our own page title)
    print("Writing preface.html")
    preface_body = section_to_html(sections["preface"], skip_count=1)
    with open(os.path.join(DOCS_DIR, "preface.html"), "w", encoding="utf-8") as f:
        f.write(page_html("Preface", preface_body, "preface"))

    # Write chapters
    for ch_num in range(15):
        key = f"chapter-{ch_num}"
        print(f"Writing {key}.html")
        if key not in sections:
            print(f"  WARNING: {key} not found in sections!")
            continue

        ch_elements = sections[key]

        # Extract subtitle: first non-empty, non-heading paragraph after the heading
        subtitle = None
        for etype, obj in ch_elements:
            if etype == "paragraph" and not obj.style.name.startswith("Heading"):
                if obj.text.strip():
                    subtitle = obj.text.strip()
                    break

        # Skip 2 paragraphs: Heading 1 ("Chapter N") + subtitle
        body = section_to_html(ch_elements, skip_count=2)

        part_num, part_name = PARTS[ch_num]
        title = f"Chapter {ch_num}: {CHAPTER_TITLES[ch_num]}"

        with open(os.path.join(DOCS_DIR, f"{key}.html"), "w", encoding="utf-8") as f:
            f.write(page_html(title, body, key, subtitle=subtitle))

    # Write appendices
    # Each appendix starts with "APPENDIX X" then the subtitle — skip both
    for app_key, app_title in [("appendix-a", "Appendix A: Mathematical Constants"),
                                ("appendix-b", "Appendix B: Notation Reference")]:
        print(f"Writing {app_key}.html")
        if app_key not in sections:
            print(f"  WARNING: {app_key} not found!")
            continue
        app_body = section_to_html(sections[app_key], skip_count=2)

        with open(os.path.join(DOCS_DIR, f"{app_key}.html"), "w", encoding="utf-8") as f:
            f.write(page_html(app_title, app_body, app_key))

    # Copy .docx for download
    dest = os.path.join(DOCS_DIR, "Ocean_From_Motion.docx")
    shutil.copy2(DOCX_PATH, dest)
    print(f"Copied .docx to {dest}")

    print("Done! Open docs/index.html in a browser to preview.")


if __name__ == "__main__":
    main()
