import { App, Notice, PluginSettingTab, Setting, SettingDefinitionItem, TFile, requireApiVersion } from "obsidian";
import type DocxStylerPlugin from "./main";
import { LicenseCache, describeLicense, normalizeLicense } from "./license";
import { GUMROAD_PRODUCT_URL, PRO_PRICE_LABEL, proIsOnSale } from "./product";

export type PageSize = "A4" | "Letter";

export interface DocxStylerSettings {
	outputFolder: string;
	fontFamily: string;
	fontSize: number;
	pageSize: PageSize;
	includeTitle: boolean;
	stripFrontmatter: boolean;
	// Pro
	licenseKey: string;
	license: LicenseCache | null;
	templatePath: string;
	pageNumbers: boolean;
}

export const DEFAULT_SETTINGS: DocxStylerSettings = {
	outputFolder: "",
	fontFamily: "Calibri",
	fontSize: 11,
	pageSize: "A4",
	includeTitle: true,
	stripFrontmatter: true,
	licenseKey: "",
	license: null,
	templatePath: "",
	pageNumbers: false,
};

/**
 * Builds settings from saved data. Unknown or old fields are dropped and a
 * license cache of any other shape (for example from an older version) counts as no license.
 */
export function migrateSettings(raw: unknown): DocxStylerSettings {
	const saved = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
	const out: DocxStylerSettings = { ...DEFAULT_SETTINGS };
	for (const key of Object.keys(DEFAULT_SETTINGS) as (keyof DocxStylerSettings)[]) {
		if (key === "license" || !(key in saved)) continue;
		const value = saved[key];
		const fallback = DEFAULT_SETTINGS[key];
		if (typeof value === typeof fallback) Object.assign(out, { [key]: cleanSetting(key, value) });
	}
	out.license = normalizeLicense(saved.license);
	return out;
}

type EditableKey = Exclude<keyof DocxStylerSettings, "license">;

/** Cleans a raw control value into the stored setting value. */
export function cleanSetting(key: EditableKey, value: unknown): DocxStylerSettings[EditableKey] {
	switch (key) {
		case "includeTitle":
		case "stripFrontmatter":
		case "pageNumbers":
			return value === true;
		case "fontSize": {
			const n = Number(value);
			return Number.isFinite(n) ? Math.min(16, Math.max(8, Math.round(n))) : DEFAULT_SETTINGS.fontSize;
		}
		case "pageSize":
			return value === "Letter" ? "Letter" : "A4";
		case "fontFamily":
			return (typeof value === "string" ? value.trim() : "") || DEFAULT_SETTINGS.fontFamily;
		default:
			return typeof value === "string" ? value.trim() : "";
	}
}

const TEXT = {
	outputFolder: ["Output folder", "Vault folder for exported files. Leave empty to save next to the note."],
	fontFamily: ["Font", "Body font used when no Word template is set."],
	fontSize: ["Font size", "Body text size in points."],
	pageSize: ["Page size", ""],
	includeTitle: ["Add note title", "Start the document with the note name as its title."],
	stripFrontmatter: ["Leave out properties", "Do not export the note's properties (frontmatter)."],
	licenseKey: ["License key", "Only the activate and check license buttons send the key to Gumroad's license API. Nothing else is sent, and nothing is sent on its own."],
	templatePath: ["Word template", "Pro. A .docx file in your vault whose styles are used for every export."],
	pageNumbers: ["Page numbers", "Pro. Add a page number footer."],
} as const;

export class DocxStylerSettingTab extends PluginSettingTab {
	plugin: DocxStylerPlugin;

	constructor(app: App, plugin: DocxStylerPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	private licenseText(): string {
		return proIsOnSale() ? describeLicense(this.plugin.settings.license) : "Pro is not on sale right now. Everything above is free and stays free.";
	}

	/** Declarative settings for Obsidian 1.13 and later (rendered and searchable by Obsidian). */
	getSettingDefinitions(): SettingDefinitionItem[] {
		return [
			{
				type: "group",
				items: [
					{ name: TEXT.outputFolder[0], desc: TEXT.outputFolder[1], control: { type: "folder", key: "outputFolder", placeholder: "Exports" } },
					{ name: TEXT.fontFamily[0], desc: TEXT.fontFamily[1], control: { type: "text", key: "fontFamily", placeholder: DEFAULT_SETTINGS.fontFamily } },
					{ name: TEXT.fontSize[0], desc: TEXT.fontSize[1], control: { type: "slider", key: "fontSize", min: 8, max: 16, step: 1 } },
					{ name: TEXT.pageSize[0], control: { type: "dropdown", key: "pageSize", options: { A4: "A4", Letter: "Letter" } } },
					{ name: TEXT.includeTitle[0], desc: TEXT.includeTitle[1], control: { type: "toggle", key: "includeTitle" } },
					{ name: TEXT.stripFrontmatter[0], desc: TEXT.stripFrontmatter[1], control: { type: "toggle", key: "stripFrontmatter" } },
				],
			},
			{
				type: "group",
				heading: "Pro",
				items: [
					{ name: "License", desc: this.licenseText(), render: (setting) => this.addBuyButton(setting) },
					{ name: TEXT.licenseKey[0], desc: TEXT.licenseKey[1], render: (setting) => this.addLicenseControls(setting) },
					{
						name: TEXT.templatePath[0],
						desc: TEXT.templatePath[1],
						control: { type: "file", key: "templatePath", placeholder: "Templates/House style.docx", filter: (f: TFile) => f.extension === "docx" },
					},
					{ name: TEXT.pageNumbers[0], desc: TEXT.pageNumbers[1], control: { type: "toggle", key: "pageNumbers" } },
				],
			},
		];
	}

	getControlValue(key: string): unknown {
		return this.plugin.settings[key as keyof DocxStylerSettings];
	}

	async setControlValue(key: string, value: unknown): Promise<void> {
		const k = key as EditableKey;
		const before = this.plugin.settings[k];
		Object.assign(this.plugin.settings, { [k]: cleanSetting(k, value) });
		// A different key needs its own activation.
		if (k === "licenseKey" && this.plugin.settings.licenseKey !== before) this.plugin.settings.license = null;
		await this.plugin.saveSettings();
	}

	/** Fallback for Obsidian older than 1.13, which does not know getSettingDefinitions(). */
	display(): void {
		this.renderLegacy();
	}

	private refresh(): void {
		if (requireApiVersion("1.13.0")) this.update();
		else this.renderLegacy();
	}

	private addBuyButton(setting: Setting): void {
		if (!proIsOnSale()) return;
		setting.addButton((b) =>
			b.setButtonText(`Buy Pro (${PRO_PRICE_LABEL})`).onClick(() => {
				window.open(GUMROAD_PRODUCT_URL);
			}),
		);
	}

	private addLicenseControls(setting: Setting): void {
		setting
			.addText((t) =>
				t
					.setPlaceholder("Paste your key")
					.setValue(this.plugin.settings.licenseKey)
					.onChange(async (v) => this.setControlValue("licenseKey", v)),
			)
			.addButton((b) =>
				b.setButtonText("Activate").onClick(async () => {
					new Notice(await this.plugin.updateLicense("activate"));
					this.refresh();
				}),
			)
			.addButton((b) =>
				b.setButtonText("Check license").onClick(async () => {
					new Notice(await this.plugin.updateLicense("check"));
					this.refresh();
				}),
			);
	}

	private renderLegacy(): void {
		const { containerEl } = this;
		containerEl.empty();
		const s = this.plugin.settings;
		const row = (key: keyof typeof TEXT) => new Setting(containerEl).setName(TEXT[key][0]).setDesc(TEXT[key][1]);
		const save = (key: EditableKey) => async (v: unknown) => this.setControlValue(key, v);

		row("outputFolder").addText((t) => t.setPlaceholder("Exports").setValue(s.outputFolder).onChange(save("outputFolder")));
		row("fontFamily").addText((t) => t.setValue(s.fontFamily).onChange(save("fontFamily")));
		row("fontSize").addSlider((sl) => sl.setLimits(8, 16, 1).setValue(s.fontSize).onChange(save("fontSize")));
		row("pageSize").addDropdown((d) => d.addOption("A4", "A4").addOption("Letter", "Letter").setValue(s.pageSize).onChange(save("pageSize")));
		row("includeTitle").addToggle((t) => t.setValue(s.includeTitle).onChange(save("includeTitle")));
		row("stripFrontmatter").addToggle((t) => t.setValue(s.stripFrontmatter).onChange(save("stripFrontmatter")));

		new Setting(containerEl).setName("Pro").setHeading();
		this.addBuyButton(new Setting(containerEl).setName("License").setDesc(this.licenseText()));
		this.addLicenseControls(row("licenseKey"));
		row("templatePath").addText((t) =>
			t.setPlaceholder("Templates/House style.docx").setValue(s.templatePath).onChange(save("templatePath")),
		);
		row("pageNumbers").addToggle((t) => t.setValue(s.pageNumbers).onChange(save("pageNumbers")));
	}
}
