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
