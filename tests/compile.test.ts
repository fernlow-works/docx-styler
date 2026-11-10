import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { naturalCompare, notesInOrder } from "../src/compile";
import { buildManuscript, buildDocx } from "../src/docx-builder";
import { prepare } from "../src/markdown";

const opts = { title: "Book", fontFamily: "Calibri", fontSize: 11, pageSize: "A4" as const, templateStylesXml: null, pageNumbers: false };
const noAssets = { load: async () => null };
const note = (title: string, body: string) => ({ title, prepared: prepare(body, true), assets: noAssets });

describe("notesInOrder", () => {
	it("sorts by path in natural order and keeps only notes", () => {
		const files = ["B/Chapter 10.md", "B/Chapter 2.md", "A/z.md", "B/pic.png", "B/chapter 1.md", "a2.md", "a10.md"].map((path) => ({ path, extension: path.split(".").pop() ?? "" }));
		expect(notesInOrder(files).map((f) => f.path)).toEqual(["A/z.md", "a2.md", "a10.md", "B/chapter 1.md", "B/Chapter 2.md", "B/Chapter 10.md"]);
	});
	it("is deterministic for names that differ only by case", () => {
		expect(naturalCompare("a.md", "A.md")).not.toBe(0);
	});
});

describe("buildManuscript", () => {
	it("joins notes in order with chapter headings, shifted note headings and isolated footnotes", async () => {
		const data = await buildManuscript([note("One", "# Inner\nText[^1]\n[^1]: First note."), note("Two", "Other[^1]\n[^1]: Second note.")], opts);
		const zip = await JSZip.loadAsync(data);
		const doc = await zip.file("word/document.xml")!.async("string");
		const foot = await zip.file("word/footnotes.xml")!.async("string");
		expect(doc.indexOf("One")).toBeLessThan(doc.indexOf("Two"));
		expect(doc).toContain('w:val="Heading1"');
		expect(doc).toContain('w:val="Heading2"');
		expect(doc).toContain("<w:pageBreakBefore");
		expect(foot).toContain("First note.");
		expect(foot).toContain("Second note.");
		expect((doc.match(/<w:footnoteReference/g) ?? []).length).toBe(2);
	});
	it("uses the template styles and page number footer when asked", async () => {
		const styles = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>';
		const data = await buildManuscript([note("One", "a"), note("Two", "b")], { ...opts, templateStylesXml: styles, pageNumbers: true });
		const zip = await JSZip.loadAsync(data);
		expect(await zip.file("word/styles.xml")!.async("string")).toContain('w:styleId="Heading1"');
		expect(await zip.file("word/footer1.xml")!.async("string")).toContain("PAGE");
	});
	it("does not break the first page when there is no title", async () => {
		const data = await buildManuscript([note("One", "a")], { ...opts, title: null });
		const doc = await (await JSZip.loadAsync(data)).file("word/document.xml")!.async("string");
		expect(doc).not.toContain("<w:pageBreakBefore");
	});
	it("refuses an empty list", async () => {
		await expect(buildManuscript([], opts)).rejects.toThrow("No notes");
	});
	it("leaves single note export unchanged", async () => {
		const data = await buildDocx(prepare("# H\nbody", true), opts, noAssets);
		const doc = await (await JSZip.loadAsync(data)).file("word/document.xml")!.async("string");
		expect(doc).toContain('w:val="Heading1"');
		expect(doc).not.toContain("<w:pageBreakBefore");
	});
});

describe("collectNotes", () => {
	it("walks subfolders and returns notes in natural order", async () => {
		const { TFile, TFolder } = await import("obsidian");
		const { collectNotes } = await import("../src/compile");
		const mk = (path: string) => Object.assign(new TFile(), { path, extension: path.split(".").pop() });
		const sub = Object.assign(new TFolder(), { path: "Book/Part", children: [mk("Book/Part/b.md"), mk("Book/Part/a.png")] });
		const root = Object.assign(new TFolder(), { path: "Book", children: [mk("Book/10 end.md"), sub, mk("Book/2 start.md")] });
		expect(collectNotes(root).map((f) => f.path)).toEqual(["Book/2 start.md", "Book/10 end.md", "Book/Part/b.md"]);
	});
});
