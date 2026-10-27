import { Menu, Notice, Plugin, TAbstractFile, TFile, TFolder, normalizePath } from "obsidian";
import JSZip from "jszip";
import { buildDocx } from "./docx-builder";
import { prepare } from "./markdown";
import { cacheGrantsPro, verifyWithGumroad } from "./license";
import { ACTIVATION_LIMIT, proIsOnSale } from "./product";
import { DEFAULT_SETTINGS, DocxStylerSettings, DocxStylerSettingTab, migrateSettings } from "./settings";

export default class DocxStylerPlugin extends Plugin {
	settings: DocxStylerSettings = DEFAULT_SETTINGS;

	async onload(): Promise<void> {
		await this.loadSettings();

		this.addRibbonIcon("file-output", "Export note to Word", () => {
			const file = this.app.workspace.getActiveFile();
			if (file) void this.exportFile(file);
			else new Notice("Open a note to export it.");
		});

		this.addCommand({
			id: "export-current-note",
			name: "Export current note to Word (.docx)",
			checkCallback: (checking) => {
				const file = this.app.workspace.getActiveFile();
				if (!file || file.extension !== "md") return false;
				if (!checking) void this.exportFile(file);
				return true;
			},
		});

		this.registerEvent(
			this.app.workspace.on("file-menu", (menu: Menu, file: TAbstractFile) => {
				if (!(file instanceof TFile) || file.extension !== "md") return;
				menu.addItem((item) =>
					item
						.setTitle("Export to Word (.docx)")
						.setIcon("file-output")
						.onClick(() => void this.exportFile(file)),
				);
			}),
		);

		this.addSettingTab(new DocxStylerSettingTab(this.app, this));
	}

	async loadSettings(): Promise<void> {
		this.settings = migrateSettings(await this.loadData());
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}

	/**
	 * Runs only from the Activate and Check license buttons. Nothing calls Gumroad on its own.
	 * Activate counts this device at Gumroad (increment_uses_count=true), Check license does not.
	 */
	async updateLicense(mode: "activate" | "check"): Promise<string> {
		const s = this.settings;
		if (!proIsOnSale()) return "Pro is not on sale yet.";
		if (!s.licenseKey) return "Enter a license key first.";
		// A device that already activated this key re-checks it instead of using another device slot.
		const alreadyActive = s.license?.status === "active";
		const activatedBefore = s.license !== null && s.license.status !== "invalid" && s.license.status !== "over-limit";
		if (mode === "check" && !activatedBefore) return "Click activate first.";
		const result = await verifyWithGumroad(s.licenseKey, mode === "activate" && !alreadyActive);
		if (result === "not-on-sale") return "Pro is not on sale yet.";
		if (result === "offline") return "Could not reach Gumroad. Nothing changed.";
		s.license = result;
		await this.saveSettings();
		switch (result.status) {
			case "active":
				return "Pro is active. Thank you!";
			case "over-limit":
				return `This key is already active on ${ACTIVATION_LIMIT} devices. Open an issue on GitHub if you need to move it.`;
			case "invalid":
				return "Gumroad did not accept this key.";
			default:
				return "This key is not active.";
		}
	}

	private proActive(): boolean {
		return cacheGrantsPro(this.settings.license);
	}

	private async templateStyles(): Promise<string | null> {
		const path = this.settings.templatePath;
		if (!path) return null;
		const file = this.app.vault.getAbstractFileByPath(normalizePath(path));
		if (!(file instanceof TFile)) {
			new Notice("Word template not found. Exporting with default styles.");
			return null;
		}
		const zip = await JSZip.loadAsync(await this.app.vault.readBinary(file));
		return (await zip.file("word/styles.xml")?.async("string")) ?? null;
	}

	async exportFile(file: TFile): Promise<void> {
		const notice = new Notice("Exporting to Word...", 0);
		try {
			const s = this.settings;
			const wantsPro = Boolean(s.templatePath) || s.pageNumbers;
			const pro = wantsPro && this.proActive();
			if (wantsPro && !pro) new Notice("Word template and page numbers are Pro features. Exporting without them.");

			const prepared = prepare(await this.app.vault.cachedRead(file), s.stripFrontmatter);
			const data = await buildDocx(
				prepared,
				{
					title: s.includeTitle ? file.basename : null,
					fontFamily: s.fontFamily,
					fontSize: s.fontSize,
					pageSize: s.pageSize,
					templateStylesXml: pro ? await this.templateStyles() : null,
					pageNumbers: pro && s.pageNumbers,
				},
				{
					load: async (target) => {
						const dest = this.app.metadataCache.getFirstLinkpathDest(target, file.path);
						return dest ? new Uint8Array(await this.app.vault.readBinary(dest)) : null;
					},
				},
			);
			const outPath = await this.outputPath(file);
			const existing = this.app.vault.getAbstractFileByPath(outPath);
			if (existing instanceof TFile) await this.app.vault.modifyBinary(existing, data);
			else await this.app.vault.createBinary(outPath, data);
			new Notice(`Saved ${outPath}`);
		} catch (e) {
			new Notice(`Export failed: ${e instanceof Error ? e.message : String(e)}`);
		} finally {
			notice.hide();
		}
	}

	private async outputPath(file: TFile): Promise<string> {
		const folder = this.settings.outputFolder ? normalizePath(this.settings.outputFolder) : (file.parent?.path ?? "");
		if (folder && folder !== "/" && !(this.app.vault.getAbstractFileByPath(folder) instanceof TFolder)) {
			await this.app.vault.createFolder(folder);
		}
		return normalizePath(`${folder && folder !== "/" ? folder + "/" : ""}${file.basename}.docx`);
	}
}
