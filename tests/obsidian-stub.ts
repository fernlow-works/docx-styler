// Test stand-in for the "obsidian" package, which ships types only.
export const requestUrl = (): never => {
	throw new Error("network is not used in tests");
};
export class PluginSettingTab {}
export class Setting {}
export class Notice {}
export class TFile {}
export const requireApiVersion = (): boolean => true;
