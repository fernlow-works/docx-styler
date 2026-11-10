"""Checks a folder-compiled .docx: chapter order, heading styles, footnotes and page numbers. Prints JSON, exits 1 on failure."""
import json, sys, zipfile
import docx

path = sys.argv[1]
d = docx.Document(path)
h1 = [p.text for p in d.paragraphs if p.style.name == "Heading 1"]
h2 = [p.text for p in d.paragraphs if p.style.name == "Heading 2"]
h3 = [p.text for p in d.paragraphs if p.style.name == "Heading 3"]
with zipfile.ZipFile(path) as z:
    names = z.namelist()
    doc_xml = z.read("word/document.xml").decode()
    foot = z.read("word/footnotes.xml").decode() if "word/footnotes.xml" in names else ""
    footer = "".join(z.read(n).decode() for n in names if n.startswith("word/footer"))

checks = {
    "title_is_folder_name": any(p.style.name == "Title" and p.text == "Book" for p in d.paragraphs),
    "chapters_in_natural_order": h1 == ["Chapter 1", "Chapter 2", "Chapter 10", "Appendix"],
    "note_headings_shifted_down_one_level": h2 == ["Opening inner heading"] and h3 == ["Deeper heading two"],
    "chapters_start_on_new_page": doc_xml.count("<w:pageBreakBefore") == 4,
    "both_footnotes_kept_apart": "Footnote from chapter one." in foot and "Footnote from chapter ten." in foot,
    "page_numbers": "PAGE" in footer,
    "frontmatter_removed": "tags:" not in "\n".join(p.text for p in d.paragraphs),
}
print(json.dumps({"file": path, "h1": h1, "checks": checks}))
sys.exit(0 if all(checks.values()) else 1)
