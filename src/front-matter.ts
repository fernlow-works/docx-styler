// Pro front matter: an optional cover page and table of contents before the body.
// Pure functions, so the rules can be tested without Obsidian.

export interface CoverOptions {
	title: string;
	subtitle: string | null;
	author: string | null;
}

export interface TocOptions {
	/** Deepest heading level listed (1 to 6). */
	levels: number;
	/** Text above the table, in the TOC Heading style. */
	title: string;
}

export interface FrontMatterSettings {
	coverPage: boolean;
	coverSubtitle: string;
	coverAuthor: string;
	tableOfContents: boolean;
	tocLevels: number;
	tocTitle: string;
}

export interface FrontMatter {
	cover: CoverOptions | null;
	toc: TocOptions | null;
}

type Properties = Record<string, unknown> | null | undefined;

/** A property as one line of text: strings are trimmed, lists are joined with commas, anything else is ignored. */
export function propertyText(props: Properties, key: string): string | null {
	const value = props?.[key];
	if (typeof value === "string") return value.trim() || null;
	if (typeof value === "number") return String(value);
	if (Array.isArray(value)) {
		const parts = value.filter((v): v is string | number => typeof v === "string" || typeof v === "number").map((v) => String(v).trim()).filter(Boolean);
		return parts.length ? parts.join(", ") : null;
	}
	return null;
}

/** A per-note on/off switch such as `docx-cover: false`. Anything that is not clearly true or false keeps the setting. */
export function propertySwitch(props: Properties, key: string, fallback: boolean): boolean {
	const value = props?.[key];
	if (value === true || value === "true") return true;
	if (value === false || value === "false") return false;
	return fallback;
}

/**
 * Works out the cover page and table of contents for one export.
 * `name` is the note name, or the folder name for a compiled folder.
 * `props` are the note's properties, or the folder note's for a compiled folder.
 */
export function resolveFrontMatter(name: string, props: Properties, s: FrontMatterSettings): FrontMatter {
	const cover = propertySwitch(props, "docx-cover", s.coverPage)
		? {
				title: propertyText(props, "title") ?? name,
				subtitle: propertyText(props, "subtitle") ?? (s.coverSubtitle.trim() || null),
				author: propertyText(props, "author") ?? (s.coverAuthor.trim() || null),
			}
		: null;
	const toc = propertySwitch(props, "docx-toc", s.tableOfContents) ? { levels: s.tocLevels, title: s.tocTitle.trim() || "Contents" } : null;
	return { cover, toc };
}

/** True when an export asks for the cover page or table of contents, both Pro. */
export function wantsFrontMatter(front: FrontMatter): boolean {
	return front.cover !== null || front.toc !== null;
}
