import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, cleanSetting, migrateSettings } from "../src/settings";

describe("cleanSetting", () => {
	it("trims text and falls back to Calibri for an empty font", () => {
		expect(cleanSetting("outputFolder", "  Exports  ")).toBe("Exports");
		expect(cleanSetting("fontFamily", "   ")).toBe("Calibri");
	});
	it("clamps and rounds the font size", () => {
		expect(cleanSetting("fontSize", 30)).toBe(16);
		expect(cleanSetting("fontSize", 2)).toBe(8);
		expect(cleanSetting("fontSize", 10.6)).toBe(11);
	});
	it("only allows A4 or Letter", () => {
		expect(cleanSetting("pageSize", "Letter")).toBe("Letter");
		expect(cleanSetting("pageSize", "Tabloid")).toBe("A4");
	});
	it("coerces toggles to booleans", () => {
		expect(cleanSetting("includeTitle", true)).toBe(true);
		expect(cleanSetting("pageNumbers", "yes")).toBe(false);
		expect(typeof DEFAULT_SETTINGS.stripFrontmatter).toBe("boolean");
	});
});

describe("migrateSettings", () => {
	it("loads empty data as defaults", () => {
		expect(migrateSettings(null)).toEqual(DEFAULT_SETTINGS);
		expect(migrateSettings(undefined)).toEqual(DEFAULT_SETTINGS);
	});
	it("keeps saved settings but drops an older license cache shape", () => {
		const old = {
			outputFolder: "Exports",
			fontSize: 12,
			licenseKey: "KEY-1",
			license: { status: "granted", expiresAt: null, activationId: "act-1", checkedAt: 1 },
		};
		const s = migrateSettings(old);
		expect(s.outputFolder).toBe("Exports");
		expect(s.fontSize).toBe(12);
		expect(s.licenseKey).toBe("KEY-1");
		expect(s.license).toBeNull();
	});
	it("keeps a current license cache and ignores wrong types and unknown fields", () => {
		const s = migrateSettings({ fontSize: "huge", unknown: 1, license: { status: "active", uses: 1, checkedAt: 9 } });
		expect(s.fontSize).toBe(DEFAULT_SETTINGS.fontSize);
		expect("unknown" in s).toBe(false);
		expect(s.license).toEqual({ status: "active", uses: 1, checkedAt: 9 });
	});
});

describe("cover page and table of contents settings", () => {
	it("are off by default with levels 1 to 3 and the title Contents", () => {
		expect(DEFAULT_SETTINGS.coverPage).toBe(false);
		expect(DEFAULT_SETTINGS.tableOfContents).toBe(false);
		expect(DEFAULT_SETTINGS.tocLevels).toBe(3);
		expect(DEFAULT_SETTINGS.tocTitle).toBe("Contents");
	});
	it("clamp the levels to 1 to 6 and fall back to Contents for an empty title", () => {
		expect(cleanSetting("tocLevels", 0)).toBe(1);
		expect(cleanSetting("tocLevels", 9)).toBe(6);
		expect(cleanSetting("tocLevels", "x")).toBe(3);
		expect(cleanSetting("tocTitle", "  ")).toBe("Contents");
		expect(cleanSetting("coverPage", "yes")).toBe(false);
		expect(cleanSetting("coverAuthor", "  Jane  ")).toBe("Jane");
	});
	it("load from 1.2.0 data with the new settings off", () => {
		const s = migrateSettings({ outputFolder: "Exports", pageNumbers: true, templatePath: "T.docx" });
		expect(s.coverPage).toBe(false);
		expect(s.tableOfContents).toBe(false);
		expect(s.pageNumbers).toBe(true);
		expect(s.templatePath).toBe("T.docx");
	});
});
