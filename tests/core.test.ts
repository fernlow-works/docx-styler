import { describe, expect, it } from "vitest";
import JSZip from "jszip";


import { prepare, lex, calloutHeader } from "../src/markdown";
import { readImageInfo } from "../src/image-size";
import { buildDocx } from "../src/docx-builder";
import { cacheGrantsPro, toCache } from "../src/license";

const opts = { title: "T", fontFamily: "Calibri", fontSize: 11, pageSize: "A4" as const, templateStylesXml: null, pageNumbers: false };

describe("prepare", () => {
	it("drops frontmatter and comments and collects footnotes", () => {
		const p = prepare("---\na: 1\n---\nText %%secret%% here[^n].\n[^n]: Note text.\n", true);
		expect(p.body).not.toContain("a: 1");
		expect(p.body).not.toContain("secret");
		expect(p.footnotes.get("n")).toBe("Note text.");
	});
});

describe("lex", () => {
	it("knows Obsidian inline syntax", () => {
		const types = JSON.stringify(lex("[[A|b]] ![[i.png|200]] ==h== x[^1]"));
		for (const t of ["wikilink", "embed", "highlight", "footnoteRef"]) expect(types).toContain(t);
	});
	it("reads callout headers", () => {
		expect(calloutHeader("[!warning]- Careful\nbody")).toEqual({ kind: "warning", title: "Careful" });
		expect(calloutHeader("plain")).toBeNull();
	});
});

describe("readImageInfo", () => {
	it("reads PNG size", () => {
		const b = new Uint8Array(32);
		new DataView(b.buffer).setUint32(0, 0x89504e47);
		new DataView(b.buffer).setUint32(16, 640);
		new DataView(b.buffer).setUint32(20, 480);
		expect(readImageInfo(b)).toEqual({ kind: "png", width: 640, height: 480 });
	});
});

describe("buildDocx", () => {
	it("uses real Word styles", async () => {
		const md = "# H\n\n- a\n\n1. b\n\n> q\n\n| x | y |\n|---|---|\n| 1 | 2 |\n\nf[^1]\n[^1]: foot\n";
		const buf = await buildDocx(prepare(md, true), opts, { load: async () => null });
		const zip = await JSZip.loadAsync(buf);
		const xml = (await zip.file("word/document.xml")?.async("string")) ?? "";
		for (const s of ["Title", "Heading1", "ListParagraph", "Quote"]) expect(xml).toContain(`w:val="${s}"`);
		expect(xml).toContain("<w:tbl>");
		expect(await zip.file("word/footnotes.xml")?.async("string")).toContain("foot");
	});
});

describe("license cache", () => {
	it("keeps only status, uses and check time from a Gumroad reply", () => {
		const c = toCache({ success: true, uses: 1, purchase: { email: "buyer@example.com", refunded: false, disputed: false, chargebacked: false } as never }, true, 1);
		expect(c).toEqual({ status: "active", uses: 1, checkedAt: 1 });
		expect(cacheGrantsPro(c)).toBe(true);
		expect(cacheGrantsPro({ ...c, status: "refunded" })).toBe(false);
	});
});
