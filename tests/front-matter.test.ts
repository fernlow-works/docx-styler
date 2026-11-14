import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildDocx, buildManuscript, BuildOptions } from "../src/docx-builder";
import { propertySwitch, propertyText, resolveFrontMatter, wantsFrontMatter } from "../src/front-matter";
import { prepare } from "../src/markdown";
import { DEFAULT_SETTINGS } from "../src/settings";

const opts: BuildOptions = { title: "Note", fontFamily: "Calibri", fontSize: 11, pageSize: "A4", templateStylesXml: null, pageNumbers: false };
const noAssets = { load: async () => null };
const note = (title: string, body: string) => ({ title, prepared: prepare(body, true), assets: noAssets });
const cover = { title: "My Book", subtitle: "A subtitle", author: "Jane Writer" };
const toc = { levels: 3, title: "Contents" };
const MD = "# One\ntext\n## Two\n### Three\n#### Four\nend";

async function parts(data: ArrayBuffer) {
	const zip = await JSZip.loadAsync(data);
	const read = async (n: string) => (await zip.file(n)?.async("string")) ?? "";
	return { doc: await read("word/document.xml"), styles: await read("word/styles.xml"), settings: await read("word/settings.xml"), core: await read("docProps/core.xml") };
}
/** Text of each paragraph in the table of contents' cached result. */
function tocEntries(doc: string): string[] {
	const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(doc)?.[0] ?? "";
	return [...sdt.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map((m) => [...(m[1] ?? "").matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("")).filter(Boolean);
}
const styleIds = (doc: string) => [...doc.matchAll(/<w:pStyle w:val="([^"]+)"\/>/g)].map((m) => m[1]);

describe("resolveFrontMatter", () => {
	const s = { ...DEFAULT_SETTINGS };
	it("is off by default", () => {
		const f = resolveFrontMatter("Note", null, s);
		expect(f).toEqual({ cover: null, toc: null });
		expect(wantsFrontMatter(f)).toBe(false);
	});
	it("uses the name and settings, and lets title, subtitle and author properties override them", () => {
		const on = { ...s, coverPage: true, coverSubtitle: " Draft ", coverAuthor: "Settings Author", tableOfContents: true, tocLevels: 2, tocTitle: "In this book" };
		expect(resolveFrontMatter("Note", null, on)).toEqual({ cover: { title: "Note", subtitle: "Draft", author: "Settings Author" }, toc: { levels: 2, title: "In this book" } });
		const f = resolveFrontMatter("Note", { title: "Real title", subtitle: "Sub", author: ["A One", "B Two"] }, on);
		expect(f.cover).toEqual({ title: "Real title", subtitle: "Sub", author: "A One, B Two" });
		expect(resolveFrontMatter("Note", null, { ...on, coverSubtitle: "", coverAuthor: "  " }).cover).toEqual({ title: "Note", subtitle: null, author: null });
	});
	it("honours docx-cover and docx-toc per note", () => {
		expect(resolveFrontMatter("N", { "docx-cover": true, "docx-toc": "true" }, s)).toEqual({ cover: { title: "N", subtitle: null, author: null }, toc: { levels: 3, title: "Contents" } });
		expect(resolveFrontMatter("N", { "docx-cover": false, "docx-toc": false }, { ...s, coverPage: true, tableOfContents: true })).toEqual({ cover: null, toc: null });
	});
	it("ignores property values that are not text or a clear switch", () => {
		expect(propertyText({ title: { a: 1 } }, "title")).toBeNull();
		expect(propertyText({ title: "  " }, "title")).toBeNull();
		expect(propertyText({ title: 2026 }, "title")).toBe("2026");
		expect(propertySwitch({ "docx-cover": "maybe" }, "docx-cover", true)).toBe(true);
	});
});

describe("cover page", () => {
	it("writes Title, Subtitle and Cover author paragraphs, then a page break, before the body", async () => {
		const { doc, styles, core } = await parts(await buildDocx(prepare(MD, true), { ...opts, cover }, noAssets));
		const ids = styleIds(doc);
		expect(ids.slice(0, 3)).toEqual(["Title", "Subtitle", "DocxStylerCoverAuthor"]);
		expect(doc.indexOf("Jane Writer")).toBeLessThan(doc.indexOf('<w:br w:type="page"/>'));
		expect(doc.indexOf('<w:br w:type="page"/>')).toBeLessThan(doc.indexOf(">One<"));
		// The cover title replaces the plain title, so it appears once.
		expect(doc.match(/w:val="Title"/g)?.length).toBe(1);
		expect(doc).not.toContain(">Note<");
		expect(styles).toContain('w:styleId="Subtitle"');
		expect(styles).toContain('w:styleId="DocxStylerCoverAuthor"');
		expect(core).toContain("My Book");
	});
	it("leaves out an empty subtitle and author", async () => {
		const { doc } = await parts(await buildDocx(prepare(MD, true), { ...opts, cover: { title: "T", subtitle: null, author: null } }, noAssets));
		expect(styleIds(doc)).not.toContain("Subtitle");
		expect(styleIds(doc)).not.toContain("DocxStylerCoverAuthor");
	});
});

describe("table of contents", () => {
	it("is a real TOC field over heading levels 1 to 3 with placeholder entries and updateFields on", async () => {
		const { doc, styles, settings } = await parts(await buildDocx(prepare(MD, true), { ...opts, toc }, noAssets));
		expect(doc).toContain("<w:sdt>");
		expect(doc).toMatch(/<w:instrText xml:space="preserve">TOC [^<]*\\o &quot;1-3&quot;[^<]*<\/w:instrText>/);
		expect(doc).toMatch(/TOC [^<]*\\h/);
		expect(doc).toContain('<w:fldChar w:fldCharType="separate"/>');
		expect(tocEntries(doc)).toEqual(["One", "Two", "Three"]);
		// Entries link to bookmarks on the headings.
		expect(doc).toMatch(/<w:hyperlink [^>]*w:anchor="_Toc\d+"/);
		expect(doc).toMatch(/<w:bookmarkStart w:name="_Toc\d+"/);
		expect(settings).toContain("<w:updateFields/>");
		expect(styleIds(doc)).toContain("TOCHeading");
		expect(styles).toMatch(/w:styleId="TOCHeading"[\s\S]*?<w:outlineLvl w:val="9"\/>/);
		// The table, then a page break, then the body.
		expect(doc.indexOf("</w:sdt>")).toBeLessThan(doc.indexOf('<w:br w:type="page"/>'));
	});
	it("follows the levels setting", async () => {
		const { doc } = await parts(await buildDocx(prepare(MD, true), { ...opts, toc: { levels: 1, title: "Contents" } }, noAssets));
		expect(doc).toContain("\\o &quot;1-1&quot;");
		expect(tocEntries(doc)).toEqual(["One"]);
	});
	it("shows a hint when the note has no headings", async () => {
		const { doc } = await parts(await buildDocx(prepare("just text", true), { ...opts, toc }, noAssets));
		expect(tocEntries(doc).join(" ")).toContain("No headings yet");
		expect(doc).toContain("TOC ");
	});
	it("goes on its own page after a cover", async () => {
		const { doc } = await parts(await buildDocx(prepare(MD, true), { ...opts, cover, toc }, noAssets));
		expect(doc).toMatch(/<w:pStyle w:val="TOCHeading"\/><w:pageBreakBefore\/>/);
		expect(doc.indexOf("Jane Writer")).toBeLessThan(doc.indexOf("<w:sdt>"));
		expect(doc.match(/<w:br w:type="page"\/>/g)?.length).toBe(1);
	});
});

describe("compiled folder with cover and table of contents", () => {
	it("lists chapters and their shifted headings, and starts chapter one on a new page without an extra break", async () => {
		const data = await buildManuscript([note("Chapter 1", "# Inner\ntext"), note("Chapter 2", "b")], { ...opts, title: "Book", cover: { ...cover, title: "Book" }, toc });
		const { doc } = await parts(data);
		expect(tocEntries(doc)).toEqual(["Chapter 1", "Inner", "Chapter 2"]);
		expect(doc.match(/<w:pageBreakBefore\/>/g)?.length).toBe(3);
		expect(doc).not.toContain('<w:br w:type="page"/>');
		expect(doc.indexOf("</w:sdt>")).toBeLessThan(doc.indexOf(">Chapter 1<", doc.indexOf("</w:sdt>")));
	});
});

describe("with a Word template", () => {
	const template = (extra: string) =>
		`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>${extra}</w:styles>`;
	it("keeps the template's own Subtitle, TOC Heading and TOC styles", async () => {
		const own = '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>';
		const { styles } = await parts(await buildDocx(prepare(MD, true), { ...opts, templateStylesXml: template(own), cover, toc }, noAssets));
		for (const id of ["Subtitle", "TOCHeading", "TOC1"]) expect(styles.match(new RegExp(`w:styleId="${id}"`, "g"))?.length).toBe(1);
		expect(styles).toContain('w:val="FF0000"');
		expect(styles).toContain('w:styleId="TOC2"');
		expect(styles).toContain('w:styleId="TOC3"');
	});
	it("adds the missing ones so the cover and table still look right", async () => {
		const { styles } = await parts(await buildDocx(prepare(MD, true), { ...opts, templateStylesXml: template(""), cover, toc }, noAssets));
		for (const id of ["Subtitle", "DocxStylerCoverAuthor", "TOCHeading", "TOC1", "TOC2", "TOC3"]) expect(styles).toContain(`w:styleId="${id}"`);
	});
});

describe("without the front matter", () => {
	it("writes the same document as before (no TOC field, no updateFields, no front styles)", async () => {
		const a = await parts(await buildDocx(prepare(MD, true), opts, noAssets));
		const b = await parts(await buildDocx(prepare(MD, true), { ...opts, cover: null, toc: null }, noAssets));
		expect(b.doc).toBe(a.doc);
		expect(b.styles).toBe(a.styles);
		expect(a.doc).not.toContain("<w:sdt>");
		expect(a.settings).not.toContain("updateFields");
		expect(a.styles).not.toContain("TOCHeading");
		expect(a.styles).not.toContain("DocxStylerCoverAuthor");
		expect(a.doc).not.toContain('<w:br w:type="page"/>');
	});
});
