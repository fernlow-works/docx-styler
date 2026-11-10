import { App, FuzzySuggestModal, TFolder } from "obsidian";

/** Lets the user pick a vault folder. */
export class FolderModal extends FuzzySuggestModal<TFolder> {
	constructor(
		app: App,
		private onPick: (folder: TFolder) => void,
	) {
		super(app);
		this.setPlaceholder("Choose a folder to compile into one Word file");
	}

	getItems(): TFolder[] {
		return this.app.vault.getAllLoadedFiles().filter((f): f is TFolder => f instanceof TFolder);
	}

	getItemText(folder: TFolder): string {
		return folder.path === "/" ? "/" : folder.path;
	}

	onChooseItem(folder: TFolder): void {
		this.onPick(folder);
	}
}
