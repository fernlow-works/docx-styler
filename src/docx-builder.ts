import {
	AlignmentType,
	Document,
	ExternalHyperlink,
	Footer,
	FootnoteReferenceRun,
	HeadingLevel,
	IParagraphStyleOptions,
	ICharacterStyleOptions,
	ImageRun,
	LevelFormat,
	Packer,
	PageBreak,
	PageNumber,
	Paragraph,
	ParagraphChild,
	Table,
	TableOfContents,
	TableCell,
	TableRow,
	TextRun,
	WidthType,
} from "docx";
import type { Tokens } from "marked";
import { calloutHeader, lex, Prepared } from "./markdown";
import { readImageInfo } from "./image-size";
import type { PageSize } from "./settings";
import type { CoverOptions, TocOptions } from "./front-matter";

export interface BuildOptions {
	title: string | null;
	fontFamily: string;
	fontSize: number;
	pageSize: PageSize;
	/** word/styles.xml from a user template (Pro). */
	templateStylesXml: string | null;
	/** Page number footer (Pro). */
	pageNumbers: boolean;
	/** Cover page before everything else (Pro). Replaces the plain title. */
	cover?: CoverOptions | null;
	/** Word table of contents field built from the heading styles (Pro). */
	toc?: TocOptions | null;
}

export interface Assets {
	/** Bytes of a vault file linked from the note, or null if it cannot be found. */
	load(target: string): Promise<Uint8Array | null>;
}

type Block = Paragraph | Table;
/** Body blocks plus the table of contents, which only the front matter has. */
type DocBlock = Block | TableOfContents;
interface Fmt {
	bold?: boolean;
	italics?: boolean;
	strike?: boolean;
	highlight?: boolean;
	code?: boolean;
}

const HEADINGS = [
	HeadingLevel.HEADING_1,
	HeadingLevel.HEADING_2,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_4,
	HeadingLevel.HEADING_5,
	HeadingLevel.HEADING_6,
] as const;
const NUMBERING_REF = "docx-styler-numbers";
const MAX_IMAGE_WIDTH_PX = 600;
/** Style ids the front matter uses. Word knows Title, Subtitle, TOC Heading and TOC 1 to 9 as built-in styles. */
export const FRONT_STYLES = { title: "Title", subtitle: "Subtitle", author: "DocxStylerCoverAuthor", tocHeading: "TOCHeading" } as const;
const PAGE = { A4: { width: 11906, height: 16838 }, Letter: { width: 12240, height: 15840 } };

function decodeEntities(s: string): string {
	return s
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&#39;/g, "'")
		.replace(/&amp;/g, "&");
}

/** One note of a compiled manuscript. */
export interface ManuscriptNote {
	/** Chapter title, shown as Heading 1. */
	title: string;
	prepared: Prepared;
	assets: Assets;
}

class Builder {
	private noteIndex = 0;
	private headingShift = 0;
	private footnoteNumbers = new Map<string, number>();
	private footnoteBodies: Record<string, { children: Paragraph[] }> = {};
	private listInstance = 0;
	/** Depth of every heading written, so the table of contents knows if it will have entries. */
	private headingDepths: number[] = [];

	constructor(
		private prepared: Prepared,
		private opts: BuildOptions,
		private assets: Assets,
	) {}

	/** Switches to the next note. Footnote ids only count inside their own note. */
	private useNote(index: number, prepared: Prepared, assets: Assets, headingShift: number): void {
		this.noteIndex = index;
		this.prepared = prepared;
		this.assets = assets;
		this.headingShift = headingShift;
	}

	private run(text: string, fmt: Fmt): TextRun {
		return new TextRun({
			text,
			bold: fmt.bold,
			italics: fmt.italics,
			strike: fmt.strike,
			highlight: fmt.highlight ? "yellow" : undefined,
			style: fmt.code ? "DocxStylerInlineCode" : undefined,
		});
	}

	private textRuns(text: string, fmt: Fmt): ParagraphChild[] {
		const out: ParagraphChild[] = [];
		decodeEntities(text)
			.split("\n")
			.forEach((line, i) => {
				if (i > 0) out.push(new TextRun({ text: "", break: 1 }));
				if (line) out.push(this.run(line, fmt));
			});
		return out;
	}

	private async image(target: string, width: number | null): Promise<ParagraphChild | null> {
		const bytes = await this.assets.load(target);
		if (!bytes) return null;
		const info = readImageInfo(bytes);
		if (!info || info.width === 0) return null;
		const w = Math.min(width ?? info.width, MAX_IMAGE_WIDTH_PX);
		const h = Math.round((info.height * w) / info.width);
		return new ImageRun({ type: info.kind, data: bytes, transformation: { width: w, height: h } });
	}

	private footnoteRef(rawId: string): ParagraphChild {
		const id = `${this.noteIndex}:${rawId}`;
		let n = this.footnoteNumbers.get(id);
		if (n === undefined) {
			n = this.footnoteNumbers.size + 1;
			this.footnoteNumbers.set(id, n);
			const text = this.prepared.footnotes.get(rawId) ?? "";
			this.footnoteBodies[String(n)] = { children: [new Paragraph({ children: this.textRuns(text, {}) })] };
		}
		return new FootnoteReferenceRun(n);
	}

	async inline(tokens: Tokens.Generic[] | undefined, fmt: Fmt = {}): Promise<ParagraphChild[]> {
		const out: ParagraphChild[] = [];
		for (const t of tokens ?? []) {
			switch (t.type) {
				case "text":
				case "escape":
					if (t.tokens?.length) out.push(...(await this.inline(t.tokens, fmt)));
					else out.push(...this.textRuns(String(t.text ?? ""), fmt));
					break;
				case "strong":
					out.push(...(await this.inline(t.tokens, { ...fmt, bold: true })));
					break;
				case "em":
					out.push(...(await this.inline(t.tokens, { ...fmt, italics: true })));
					break;
				case "del":
					out.push(...(await this.inline(t.tokens, { ...fmt, strike: true })));
					break;
				case "highlight":
					out.push(...(await this.inline(t.tokens, { ...fmt, highlight: true })));
					break;
				case "codespan":
					out.push(...this.textRuns(String(t.text ?? ""), { ...fmt, code: true }));
					break;
				case "br":
					out.push(new TextRun({ text: "", break: 1 }));
					break;
				case "link": {
					const href = String(t.href ?? "");
					const children = (await this.inline(t.tokens, fmt)).filter((c): c is TextRun => c instanceof TextRun);
					if (/^(https?:|mailto:)/i.test(href)) {
						const label = decodeEntities(String(t.text ?? href));
						out.push(new ExternalHyperlink({ link: href, children: [new TextRun({ text: label, style: "Hyperlink" })] }));
					} else {
						out.push(...children);
					}
					break;
				}
				case "wikilink":
					out.push(this.run(String(t.label), fmt));
					break;
				case "image":
				case "embed": {
					const target = t.type === "image" ? decodeURIComponent(String(t.href ?? "")) : String(t.target);
					// Remote images are never fetched: the only network call this plugin makes is the license check.
					const img = /^https?:/i.test(target) ? null : await this.image(target, t.type === "embed" ? (t.width as number | null) : null);
					out.push(img ?? this.run(`[${t.type === "image" ? String(t.text || target) : target}]`, { ...fmt, italics: true }));
					break;
				}
				case "footnoteRef":
					out.push(this.footnoteRef(String(t.id)));
					break;
				case "html":
					out.push(...this.textRuns(String(t.raw).replace(/<[^>]*>/g, ""), fmt));
					break;
				default:
					if (t.raw) out.push(...this.textRuns(String(t.raw), fmt));
			}
		}
		return out;
	}

	private async list(token: Tokens.Generic, level: number, out: Block[]): Promise<void> {
		const ordered = Boolean(token.ordered);
		const instance = ordered ? ++this.listInstance : 0;
		for (const item of (token.items ?? []) as Tokens.ListItem[]) {
			let first = true;
			for (const child of item.tokens as Tokens.Generic[]) {
				if (child.type === "list") {
					await this.list(child, Math.min(level + 1, 8), out);
					continue;
				}
				if (first && (child.type === "text" || child.type === "paragraph")) {
					const prefix = item.task ? [new TextRun({ text: item.checked ? "\u2612 " : "\u2610 " })] : [];
					const children = [...prefix, ...(await this.inline(child.tokens ?? [child]))];
					out.push(
						new Paragraph({
							children,
							...(ordered
								? { numbering: { reference: NUMBERING_REF, level, instance } }
								: { bullet: { level } }),
						}),
					);
					first = false;
					continue;
				}
				out.push(...(await this.blocks([child])));
			}
		}
	}

	private async quote(token: Tokens.Generic): Promise<Block[]> {
		const text = String(token.text ?? "");
		const callout = calloutHeader(text);
		const innerTokens = callout ? lex(text.split("\n").slice(1).join("\n")) : (token.tokens ?? []);
		const out: Block[] = [];
		if (callout) out.push(new Paragraph({ style: "Quote", children: [new TextRun({ text: callout.title, bold: true })] }));
		for (const t of innerTokens) {
			if (t.type === "paragraph" || t.type === "text") {
				out.push(new Paragraph({ style: "Quote", children: await this.inline(t.tokens ?? [t]) }));
			} else {
				out.push(...(await this.blocks([t])));
			}
		}
		return out;
	}

	private async table(token: Tokens.Generic): Promise<Table> {
		const header = token.header as Tokens.TableCell[];
		const rows = token.rows as Tokens.TableCell[][];
		const cell = async (c: Tokens.TableCell, bold: boolean) =>
			new TableCell({ children: [new Paragraph({ children: await this.inline(c.tokens, { bold }) })] });
		const tableRows = [new TableRow({ tableHeader: true, children: await Promise.all(header.map((c) => cell(c, true))) })];
		for (const r of rows) tableRows.push(new TableRow({ children: await Promise.all(r.map((c) => cell(c, false))) }));
		return new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
	}

	async blocks(tokens: Tokens.Generic[]): Promise<Block[]> {
		const out: Block[] = [];
		for (const t of tokens) {
			switch (t.type) {
				case "heading": {
					const depth = Math.min(Math.max(Number(t.depth), 1) + this.headingShift, 6);
					this.headingDepths.push(depth);
					out.push(new Paragraph({ heading: HEADINGS[depth - 1], children: await this.inline(t.tokens) }));
					break;
				}
				case "paragraph":
				case "text":
					out.push(new Paragraph({ children: await this.inline(t.tokens ?? [t]) }));
					break;
				case "list":
					await this.list(t, 0, out);
					break;
				case "blockquote":
					out.push(...(await this.quote(t)));
					break;
				case "code":
					for (const line of String(t.text ?? "").split("\n")) {
						out.push(new Paragraph({ style: "DocxStylerCode", children: [new TextRun({ text: line })] }));
					}
					break;
				case "table":
					out.push(await this.table(t));
					out.push(new Paragraph({ children: [] }));
					break;
				case "hr":
					out.push(new Paragraph({ thematicBreak: true, children: [] }));
					break;
				case "html":
					out.push(new Paragraph({ children: this.textRuns(String(t.raw).replace(/<[^>]*>/g, "").trim(), {}) }));
					break;
				case "space":
					break;
				default:
					if (t.raw) out.push(new Paragraph({ children: this.textRuns(String(t.raw).trim(), {}) }));
			}
		}
		return out;
	}

	private styles() {
		const template = this.opts.templateStylesXml;
		const has = (id: string) => template !== null && template.includes(`w:styleId="${id}"`);
		const paragraphStyles: IParagraphStyleOptions[] = [
			{
				id: "Quote",
				name: "Quote",
				basedOn: "Normal",
				next: "Normal",
				run: { italics: true, color: "555555" },
				paragraph: { indent: { left: 567 }, spacing: { before: 60, after: 60 } },
			},
			{
				id: "DocxStylerCode",
				name: "Code block",
				basedOn: "Normal",
				run: { font: "Consolas", size: 18 },
				paragraph: { spacing: { before: 0, after: 0 }, shading: { type: "clear", fill: "F2F2F2", color: "auto" } },
			},
			...this.frontStyles(),
		].filter((s) => !has(s.id)) as IParagraphStyleOptions[];
		const characterStyles: ICharacterStyleOptions[] = [
			{ id: "DocxStylerInlineCode", name: "Inline code", basedOn: "DefaultParagraphFont", run: { font: "Consolas", shading: { type: "clear", fill: "F2F2F2", color: "auto" } } },
		].filter((s) => !has(s.id)) as ICharacterStyleOptions[];
		return {
			...(template ? {} : { default: { document: { run: { font: this.opts.fontFamily, size: this.opts.fontSize * 2 } } } }),
			paragraphStyles,
			characterStyles,
		};
	}

	/**
	 * Fallback styles for the front matter, added only when it is used and only when the template (or the default
	 * styles) do not already define them, so a buyer's template styles the cover and the table of contents.
	 */
	private frontStyles(): IParagraphStyleOptions[] {
		const out: IParagraphStyleOptions[] = [];
		if (this.opts.cover) {
			out.push(
				{ id: FRONT_STYLES.subtitle, name: "Subtitle", basedOn: "Normal", next: "Normal", run: { color: "5A5A5A", size: 28 }, paragraph: { spacing: { after: 240 } } },
				{ id: FRONT_STYLES.author, name: "Cover author", basedOn: FRONT_STYLES.subtitle, next: "Normal", run: { size: 24, color: "404040" } },
			);
		}
		if (this.opts.toc) {
			// Outline level 9 (written as 9, "body text") keeps the TOC title itself out of the table, as in Word.
			// TOC 1 to 9 come with the default styles and give way to a template's own.
			out.push({ id: FRONT_STYLES.tocHeading, name: "TOC Heading", basedOn: "Heading1", next: "Normal", paragraph: { outlineLevel: 9 } });
		}
		return out;
	}

	/** Cover page paragraphs: title, then subtitle and author when set. */
	private coverBlocks(cover: CoverOptions): Paragraph[] {
		const out = [new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(cover.title)] })];
		if (cover.subtitle) out.push(new Paragraph({ style: FRONT_STYLES.subtitle, children: [new TextRun(cover.subtitle)] }));
		if (cover.author) out.push(new Paragraph({ style: FRONT_STYLES.author, children: [new TextRun(cover.author)] }));
		return out;
	}

	/**
	 * A real Word TOC field over the heading styles. Its cached result lists the document's headings (linked to
	 * bookmarks on them) so it is not empty before Word updates it, and updateFields in settings.xml makes Word
	 * refresh it, page numbers included, when the file opens.
	 */
	private tocBlocks(toc: TocOptions, afterCover: boolean): DocBlock[] {
		const listed = this.headingDepths.some((d) => d <= toc.levels);
		return [
			new Paragraph({ style: FRONT_STYLES.tocHeading, pageBreakBefore: afterCover, children: [new TextRun(toc.title)] }),
			new TableOfContents(toc.title, {
				hyperlink: true,
				headingStyleRange: `1-${toc.levels}`,
				hideTabAndPageNumbersInWebView: true,
				useAppliedParagraphOutlineLevel: true,
				// With no headings in range, a short hint stands in until headings are added and Word updates the table.
				...(listed ? {} : { contentChildren: [new Paragraph({ children: [new TextRun("No headings yet. Add headings, then update this table in Word.")] })] }),
			}),
		];
	}

	async build(notes?: ManuscriptNote[]): Promise<ArrayBuffer> {
		const { cover = null, toc = null } = this.opts;
		const body: Block[] = [];
		if (notes) {
			// Chapter title is Heading 1 and the note's own headings move down one level.
			const startsOnNewPage = Boolean(this.opts.title) || cover !== null || toc !== null;
			for (const [i, note] of notes.entries()) {
				this.useNote(i, note.prepared, note.assets, 1);
				this.headingDepths.push(1);
				body.push(
					new Paragraph({
						heading: HeadingLevel.HEADING_1,
						pageBreakBefore: i > 0 || startsOnNewPage,
						children: [new TextRun(note.title)],
					}),
				);
				body.push(...(await this.blocks(lex(note.prepared.body))));
			}
		} else {
			body.push(...(await this.blocks(lex(this.prepared.body))));
		}

		// Front matter goes before the body. It is built after it so the table of contents knows the headings.
		const children: DocBlock[] = [];
		if (cover) children.push(...this.coverBlocks(cover));
		else if (this.opts.title) children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(this.opts.title)] }));
		if (toc) children.push(...this.tocBlocks(toc, cover !== null));
		// A compiled folder's first chapter already starts on a new page.
		if (!notes && (cover || toc)) children.push(new Paragraph({ children: [new PageBreak()] }));
		children.push(...body);

		const doc = new Document({
			creator: "",
			title: cover?.title ?? this.opts.title ?? "",
			...(toc ? { features: { updateFields: true } } : {}),
			...(this.opts.templateStylesXml ? { externalStyles: this.opts.templateStylesXml } : {}),
			styles: this.styles(),
			footnotes: this.footnoteBodies,
			numbering: {
				config: [
					{
						reference: NUMBERING_REF,
						levels: Array.from({ length: 9 }, (_, level) => ({
							level,
							format: LevelFormat.DECIMAL,
							text: `%${level + 1}.`,
							alignment: AlignmentType.START,
							style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
						})),
					},
				],
			},
			sections: [
				{
					properties: {
						page: { size: PAGE[this.opts.pageSize], margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
					},
					footers: this.opts.pageNumbers
						? {
								default: new Footer({
									children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })],
								}),
							}
						: undefined,
					children,
				},
			],
		});
		return Packer.toArrayBuffer(doc);
	}
}

export function buildDocx(prepared: Prepared, opts: BuildOptions, assets: Assets): Promise<ArrayBuffer> {
	return new Builder(prepared, opts, assets).build();
}

/** Joins several notes, in the given order, into one document with the same styles and page setup as a single export. */
export function buildManuscript(notes: ManuscriptNote[], opts: BuildOptions): Promise<ArrayBuffer> {
	const first = notes[0];
	if (!first) return Promise.reject(new Error("No notes to compile."));
	return new Builder(first.prepared, opts, first.assets).build(notes);
}
