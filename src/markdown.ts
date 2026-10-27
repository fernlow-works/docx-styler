import { Marked, Tokens, TokenizerAndRendererExtension } from "marked";

export interface Prepared {
	body: string;
	/** Footnote id from the note mapped to its text. */
	footnotes: Map<string, string>;
}

/** Obsidian syntax that plain Markdown parsers do not know, removed or set aside before parsing. */
export function prepare(source: string, stripFrontmatter: boolean): Prepared {
	let body = source.replace(/\r\n/g, "\n");
	if (stripFrontmatter) body = body.replace(/^---\n[\s\S]*?\n---\n?/, "");
	body = body.replace(/%%[\s\S]*?%%/g, "");
	const footnotes = new Map<string, string>();
	body = body.replace(/^\[\^([^\]\s]+)\]:[ \t]+(.+)$/gm, (_m, id: string, text: string) => {
		footnotes.set(id, text.trim());
		return "";
	});
	return { body, footnotes };
}

export interface HighlightToken extends Tokens.Generic {
	type: "highlight";
	tokens: Tokens.Generic[];
}
export interface WikilinkToken extends Tokens.Generic {
	type: "wikilink";
	target: string;
	label: string;
}
export interface EmbedToken extends Tokens.Generic {
	type: "embed";
	target: string;
	width: number | null;
}
export interface FootnoteRefToken extends Tokens.Generic {
	type: "footnoteRef";
	id: string;
}

const embed: TokenizerAndRendererExtension = {
	name: "embed",
	level: "inline",
	start: (src) => src.indexOf("![["),
	tokenizer(src) {
		const m = /^!\[\[([^\]|#]+)(?:#[^\]|]*)?(?:\|([^\]]*))?\]\]/.exec(src);
		if (!m) return undefined;
		const width = m[2] && /^\d+/.test(m[2]) ? parseInt(m[2], 10) : null;
		return { type: "embed", raw: m[0], target: (m[1] ?? "").trim(), width };
	},
};

const wikilink: TokenizerAndRendererExtension = {
	name: "wikilink",
	level: "inline",
	start: (src) => src.indexOf("[["),
	tokenizer(src) {
		const m = /^\[\[([^\]|]+)(?:\|([^\]]*))?\]\]/.exec(src);
		if (!m) return undefined;
		const target = (m[1] ?? "").trim();
		const label = (m[2] ?? "").trim() || target.replace(/#\^?/, " > ");
		return { type: "wikilink", raw: m[0], target, label };
	},
};

const highlight: TokenizerAndRendererExtension = {
	name: "highlight",
	level: "inline",
	start: (src) => src.indexOf("=="),
	tokenizer(src) {
		const m = /^==([^=\n]+)==/.exec(src);
		if (!m) return undefined;
		return { type: "highlight", raw: m[0], tokens: this.lexer.inlineTokens(m[1] ?? "") };
	},
};

const footnoteRef: TokenizerAndRendererExtension = {
	name: "footnoteRef",
	level: "inline",
	start: (src) => src.indexOf("[^"),
	tokenizer(src) {
		const m = /^\[\^([^\]\s]+)\]/.exec(src);
		if (!m) return undefined;
		return { type: "footnoteRef", raw: m[0], id: m[1] ?? "" };
	},
};

const parser = new Marked({ gfm: true, extensions: [embed, wikilink, highlight, footnoteRef] });

export function lex(body: string): Tokens.Generic[] {
	return parser.lexer(body);
}

/** Obsidian callout header such as "[!note] Title" or "[!tip]- Title". */
export function calloutHeader(text: string): { kind: string; title: string } | null {
	const m = /^\[!([\w-]+)\][+-]?[ \t]*(.*)$/m.exec(text.split("\n")[0] ?? "");
	if (!m) return null;
	const kind = (m[1] ?? "note").toLowerCase();
	return { kind, title: (m[2] ?? "").trim() || kind.charAt(0).toUpperCase() + kind.slice(1) };
}
