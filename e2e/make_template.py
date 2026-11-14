"""Makes a Word template from an exported .docx by adding its own Subtitle and TOC Heading styles (dark red), as a buyer's template would. Usage: make_template.py in.docx out.docx"""
import shutil, sys, zipfile

src, dst = sys.argv[1], sys.argv[2]
OWN = (
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
    '<w:rPr><w:i/><w:color w:val="8B0000"/><w:sz w:val="32"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/>'
    '<w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:rPr><w:color w:val="8B0000"/></w:rPr></w:style>'
)
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst + ".tmp", "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/styles.xml":
            text = data.decode()
            assert 'w:styleId="Subtitle"' not in text
            data = text.replace("</w:styles>", OWN + "</w:styles>").encode()
        zout.writestr(item, data)
shutil.move(dst + ".tmp", dst)
