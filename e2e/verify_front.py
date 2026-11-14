"""Checks the Pro cover page and table of contents in an exported .docx and in its LibreOffice PDF.
Usage: verify_front.py file.docx file.pdf '<json expectations>'. Prints JSON, exits 1 on failure.
Expectations: title, subtitle, author, entries (TOC entry texts in order), body_marker (text that starts the body),
template_color (Subtitle color the buyer's template sets, or null)."""
import json, re, subprocess, sys, zipfile
import docx

path, pdf, want = sys.argv[1], sys.argv[2], json.loads(sys.argv[3])
d = docx.Document(path)
paras = [(p.style.style_id, p.text) for p in d.paragraphs]
with zipfile.ZipFile(path) as z:
    doc_xml = z.read("word/document.xml").decode()
    settings = z.read("word/settings.xml").decode()
    styles = z.read("word/styles.xml").decode()
sdt = re.search(r"<w:sdt>.*?</w:sdt>", doc_xml, re.S)
sdt_xml = sdt.group(0) if sdt else ""
entries = []
for p in re.findall(r"<w:p>(.*?)</w:p>", sdt_xml, re.S):
    t = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p))
    if t:
        entries.append(t)


def page(n):
    return subprocess.run(["pdftotext", "-layout", "-f", str(n), "-l", str(n), pdf, "-"], capture_output=True, text=True).stdout


pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
p1, p2, p3 = page(1), page(2), page(3)
first = [s for s, _ in paras[:3]]
checks = {
    "cover_title_style_first": paras[0] == ("Title", want["title"]),
    "cover_subtitle_style": ("Subtitle", want["subtitle"]) in paras[:3],
    "cover_author_style": ("DocxStylerCoverAuthor", want["author"]) in paras[:3],
    "cover_title_once": [s for s, _ in paras].count("Title") == 1,
    "toc_heading_style": any(s == "TOCHeading" for s, _ in paras),
    "toc_field_over_heading_styles": bool(re.search(r"TOC [^<]*\\o &quot;1-3&quot;", sdt_xml)) and "\\h" in sdt_xml,
    "toc_placeholder_entries": entries == want["entries"],
    "toc_entries_link_to_headings": len(re.findall(r'w:anchor="_Toc\d+"', sdt_xml)) == len(want["entries"]),
    "update_fields_on_open": "<w:updateFields/>" in settings or '<w:updateFields w:val="true"/>' in settings,
    "pdf_page1_is_cover": want["title"] in p1 and want["author"] in p1 and "Contents" not in p1 and want["body_marker"] not in p1,
    "pdf_page2_is_toc": "Contents" in p2 and all(e in p2 for e in want["entries"]) and want["body_marker"] not in p2,
    "pdf_page3_starts_body": want["body_marker"] in p3,
}
if want.get("template_color"):
    sub = re.search(r'<w:style [^>]*w:styleId="Subtitle".*?</w:style>', styles, re.S)
    checks["template_subtitle_style_kept"] = bool(sub) and want["template_color"] in sub.group(0) and styles.count('w:styleId="Subtitle"') == 1
print(json.dumps({"file": path, "pages": pages, "first_styles": first, "entries": entries, "checks": checks}))
sys.exit(0 if all(checks.values()) else 1)
