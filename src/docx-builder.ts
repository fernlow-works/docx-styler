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
	PageNumber,
	Paragraph,
	ParagraphChild,
	Table,
	TableCell,
	TableRow,
	TextRun,
	WidthType,
} from "docx";
import type { Tokens } from "marked";
import { calloutHeader, lex, Prepared } from "./markdown";
import { readImageInfo } from "./image-size";
import type { PageSize } from "./settings";

export interface BuildOptions {
	title: string | null;
	fontFamily: string;
	fontSize: number;
	pageSize: PageSize;
	/** word/styles.xml from a user template (Pro). */
	templateStylesXml: string | null;
	/** Page number footer (Pro). */
	pageNumbers: boolean;
}

export interface Assets {
	/** Bytes of a vault file linked from the note, or null if it cannot be found. */
	load(target: string): Promise<Uint8Array | null>;
}

type Block = Paragraph | Table;
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
const PAGE = { A4: { width: 11906, height: 16838 }, Letter: { width: 12240, height: 15840 } };

function decodeEntities(s: string): string {
	return s
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&#39;/g, "'")
		.replace(/&amp;/g, "&");
}

class Builder {
	private footnoteNumbers = new Map<string, number>();
	private footnoteBodies: Record<string, { children: Paragraph[] }> = {};
	private listInstance = 0;

	constructor(
		private prepared: Prepared,
		private opts: BuildOptions,
		private assets: Assets,
	) {}

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

	private footnoteRef(id: string): ParagraphChild {
		let n = this.footnoteNumbers.get(id);
		if (n === undefined) {
			n = this.footnoteNumbers.size + 1;
			this.footnoteNumbers.set(id, n);
			const text = this.prepared.footnotes.get(id) ?? "";
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
					const depth = Math.min(Math.max(Number(t.depth), 1), 6);
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

	async build(): Promise<ArrayBuffer> {
		const children: Block[] = [];
		if (this.opts.title) children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(this.opts.title)] }));
		children.push(...(await this.blocks(lex(this.prepared.body))));

		const doc = new Document({
			creator: "",
			title: this.opts.title ?? "",
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
