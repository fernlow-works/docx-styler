import { TFile, TFolder } from "obsidian";

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

/** Natural order by path, so "Chapter 2" comes before "Chapter 10". */
export function naturalCompare(a: string, b: string): number {
	return collator.compare(a, b) || (a < b ? -1 : a > b ? 1 : 0);
}

/** Every Markdown note below the folder (subfolders included), in natural path order. */
export function notesInOrder<T extends { path: string; extension: string }>(files: readonly T[]): T[] {
	return files.filter((f) => f.extension === "md").sort((a, b) => naturalCompare(a.path, b.path));
}

export function collectNotes(folder: TFolder): TFile[] {
	const out: TFile[] = [];
	const walk = (f: TFolder): void => {
		for (const child of f.children) {
			if (child instanceof TFolder) walk(child);
			else if (child instanceof TFile) out.push(child);
		}
	};
	walk(folder);
	return notesInOrder(out);
}
