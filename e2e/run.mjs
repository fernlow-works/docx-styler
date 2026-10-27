// End-to-end test in the official Obsidian Linux AppImage.
// Run with `npm run e2e` (wraps this in xvfb-run). Needs `npm run build` first.
// Work dir (AppImage cache, throwaway vault, config, results) defaults to ../e2e next to this repo.
import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const WORK = path.resolve(process.env.DOCX_STYLER_E2E_DIR ?? path.join(REPO, "..", "e2e"));
const OBSIDIAN_VERSION = process.env.OBSIDIAN_VERSION ?? "1.14.5";
const APPIMAGE_URL = `https://github.com/obsidianmd/obsidian-releases/releases/download/v${OBSIDIAN_VERSION}/Obsidian-${OBSIDIAN_VERSION}.AppImage`;
const CACHE = path.join(WORK, "cache");
const APP_DIR = path.join(CACHE, `obsidian-${OBSIDIAN_VERSION}`);
const RUN = path.join(WORK, "run");
const VAULT = path.join(RUN, "vault");
const HOME = path.join(RUN, "home");
const RESULTS = path.join(WORK, "results");
const PLUGIN_ID = "docx-styler";
const PORT = 9300 + Math.floor(Math.random() * 500);

const results = [];
let failed = 0;
function check(name, ok, detail = "") {
	results.push({ name, ok: Boolean(ok), detail });
	if (!ok) failed++;
	console.log(`${ok ? "PASS" : "FAIL"} ${name}${detail ? `  (${detail})` : ""}`);
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function ensureObsidian() {
	const bin = path.join(APP_DIR, "obsidian");
	if (fs.existsSync(bin)) return bin;
	fs.mkdirSync(CACHE, { recursive: true });
	const image = path.join(CACHE, `Obsidian-${OBSIDIAN_VERSION}.AppImage`);
	if (!fs.existsSync(image)) {
		console.log(`Downloading ${APPIMAGE_URL}`);
		const r = spawnSync("curl", ["-fsSL", "-o", image, APPIMAGE_URL], { stdio: "inherit" });
		if (r.status !== 0) throw new Error("AppImage download failed");
		fs.chmodSync(image, 0o755);
	}
	spawnSync(image, ["--appimage-extract"], { cwd: CACHE, stdio: "ignore" });
	fs.renameSync(path.join(CACHE, "squashfs-root"), APP_DIR);
	return bin;
}

function prepareVault() {
	fs.rmSync(RUN, { recursive: true, force: true });
	const pluginDir = path.join(VAULT, ".obsidian", "plugins", PLUGIN_ID);
	fs.mkdirSync(pluginDir, { recursive: true });
	for (const f of ["main.js", "manifest.json", "styles.css"]) fs.copyFileSync(path.join(REPO, f), path.join(pluginDir, f));
	for (const f of fs.readdirSync(path.join(REPO, "e2e", "fixtures"))) fs.copyFileSync(path.join(REPO, "e2e", "fixtures", f), path.join(VAULT, f));
	const cfg = path.join(HOME, ".config", "obsidian");
	fs.mkdirSync(cfg, { recursive: true });
	fs.writeFileSync(
		path.join(cfg, "obsidian.json"),
		JSON.stringify({ updateDisabled: true, vaults: { e2etestvault0001: { path: VAULT, ts: Date.now(), open: true } } }),
	);
}

async function connect() {
	for (let i = 0; i < 60; i++) {
		try {
			const browser = await chromium.connectOverCDP(`http://127.0.0.1:${PORT}`);
			for (let j = 0; j < 60; j++) {
				for (const ctx of browser.contexts())
					for (const page of ctx.pages()) {
						const ready = await page.evaluate(() => Boolean(window.app?.workspace?.layoutReady)).catch(() => false);
						if (ready) return { browser, page };
					}
				await sleep(500);
			}
			throw new Error("Obsidian window never became ready");
		} catch (e) {
			if (String(e).includes("never became ready")) throw e;
			await sleep(500);
		}
	}
	throw new Error("Could not connect to Obsidian over CDP");
}

/** Everything a plugin could leave behind: DOM nodes, window/document listeners, app event handlers, intervals. */
async function snapshot(page, cdp) {
	const counts = await page.evaluate((id) => {
		const handlers = (o) => Object.values(o?._ ?? {}).reduce((n, a) => n + (Array.isArray(a) ? a.length : 0), 0);
		return {
			domNodes: document.getElementsByTagName("*").length,
			pluginNodes: document.querySelectorAll(`[class*="${id}"], [data-plugin-id="${id}"]`).length,
			notices: document.querySelectorAll(".notice").length,
			workspaceHandlers: handlers(app.workspace),
			vaultHandlers: handlers(app.vault),
			metadataHandlers: handlers(app.metadataCache),
			commands: Object.keys(app.commands.commands).filter((c) => c.startsWith(`${id}:`)).length,
			ribbonItems: document.querySelectorAll('.side-dock-ribbon-action[aria-label="Export note to Word"]').length,
			settingTabs: app.setting.pluginTabs.filter((t) => t.id === id).length,
			loaded: Boolean(app.plugins.plugins[id]),
			liveIntervals: window.__e2eIntervals ? [...window.__e2eIntervals.values()].filter((s) => s.includes(`plugin:${id}`)).length : 0,
		};
	}, PLUGIN_ID);
	for (const expr of ["window", "document", "document.body"]) {
		const { result } = await cdp.send("Runtime.evaluate", { expression: expr });
		const { listeners } = await cdp.send("DOMDebugger.getEventListeners", { objectId: result.objectId });
		counts[`listeners:${expr}`] = listeners.length;
	}
	return counts;
}

async function waitForFile(rel, sinceMs, timeout = 15000) {
	const abs = path.join(VAULT, rel);
	const end = Date.now() + timeout;
	while (Date.now() < end) {
		if (fs.existsSync(abs) && fs.statSync(abs).mtimeMs >= sinceMs && fs.statSync(abs).size > 0) {
			await sleep(300);
			return abs;
		}
		await sleep(200);
	}
	return null;
}

async function waitNoticesGone(page) {
	await page.waitForFunction(() => document.querySelectorAll(".notice").length === 0, null, { timeout: 20000 }).catch(() => {});
}

function verifyDocx(file, label) {
	const keep = path.join(RESULTS, `${label}.docx`);
	fs.copyFileSync(file, keep);
	const py = spawnSync("python3", [path.join(REPO, "e2e", "verify_docx.py"), keep], { encoding: "utf8" });
	let parsed = null;
	try {
		parsed = JSON.parse(py.stdout);
	} catch {
		/* reported below */
	}
	const bad = parsed ? Object.entries(parsed.checks).filter(([, v]) => !v).map(([k]) => k) : ["unparsable"];
	check(`${label}: python-docx structure checks`, py.status === 0, bad.length ? `failed: ${bad.join(", ")} ${py.stderr}` : `${Object.keys(parsed.checks).length} checks`);
	const lo = spawnSync("soffice", ["--headless", `-env:UserInstallation=file://${path.join(RUN, "lo-profile")}`, "--convert-to", "pdf", "--outdir", RESULTS, keep], { encoding: "utf8", timeout: 120000 });
	const pdf = keep.replace(/\.docx$/, ".pdf");
	const text = fs.existsSync(pdf) ? spawnSync("pdftotext", [pdf, "-"], { encoding: "utf8" }).stdout : "";
	check(`${label}: LibreOffice opens and converts to PDF`, lo.status === 0 && text.includes("Report title") && text.includes("Section three") && text.includes("b1"), `${text.length} chars of text`);
}

async function main() {
	if (!fs.existsSync(path.join(REPO, "main.js"))) throw new Error("Run npm run build first");
	fs.rmSync(RESULTS, { recursive: true, force: true });
	fs.mkdirSync(RESULTS, { recursive: true });
	const bin = ensureObsidian();
	prepareVault();

	const child = spawn(bin, ["--no-sandbox", `--remote-debugging-port=${PORT}`, "--disable-gpu"], {
		env: { ...process.env, HOME, XDG_CONFIG_HOME: path.join(HOME, ".config") },
		detached: true,
		stdio: ["ignore", "pipe", "pipe"],
	});
	const appLog = fs.createWriteStream(path.join(RESULTS, "obsidian.log"));
	child.stdout.pipe(appLog);
	child.stderr.pipe(appLog);
	const consoleErrors = [];
	let browser;
	try {
		const conn = await connect();
		browser = conn.browser;
		const page = conn.page;
		page.on("console", (m) => {
			if (m.type() === "error" || m.type() === "warning") consoleErrors.push(`${m.type()}: ${m.text()}`);
		});
		page.on("pageerror", (e) => consoleErrors.push(`pageerror: ${e.message}`));
		const cdp = await page.context().newCDPSession(page);
		const version = await page.evaluate(() => navigator.userAgent.match(/obsidian\/([\d.]+)/i)?.[1] ?? "unknown");
		check("Obsidian started from the AppImage", true, `app ${version}, installer ${OBSIDIAN_VERSION}`);

		// Track intervals with the creating stack so plugin-made ones can be told apart.
		await page.evaluate(() => {
			const live = new Map();
			const set = window.setInterval.bind(window);
			const clear = window.clearInterval.bind(window);
			window.setInterval = (...a) => {
				const id = set(...a);
				live.set(id, new Error().stack ?? "");
				return id;
			};
			window.clearInterval = (id) => {
				live.delete(id);
				return clear(id);
			};
			window.__e2eIntervals = live;
		});

		// A fresh vault that ships a plugin folder opens with the "Do you trust the author" prompt.
		// Stay in Restricted Mode here, then turn community plugins on below the way a user would.
		const trust = page.locator(".modal-container.mod-confirmation button", { hasText: "Restricted Mode" });
		await trust.waitFor({ timeout: 10000 }).then(() => trust.click()).catch(() => {});
		await page.waitForFunction(() => !document.querySelector(".modal-container"), null, { timeout: 5000 });

		await page.evaluate(() => app.workspace.getLeaf(true).openFile(app.vault.getFileByPath("Sample.md")));
		await sleep(1000);
		await waitNoticesGone(page);
		const before = await snapshot(page, cdp);

		// Enable
		await page.evaluate(async (id) => {
			await app.plugins.setEnable(true);
			await app.plugins.loadManifests();
			await app.plugins.enablePluginAndSave(id);
		}, PLUGIN_ID);
		await sleep(500);
		const enabled = await snapshot(page, cdp);
		check("Plugin enables", enabled.loaded && enabled.commands === 1 && enabled.ribbonItems === 1 && enabled.settingTabs === 1, JSON.stringify({ commands: enabled.commands, ribbon: enabled.ribbonItems, tabs: enabled.settingTabs }));

		// Settings tab renders through the 1.13+ declarative API and is searchable.
		const settings = await page.evaluate(async (id) => {
			app.setting.open();
			app.setting.openTabById(id);
			await new Promise((r) => setTimeout(r, 800));
			const text = app.setting.modalEl?.textContent ?? document.body.textContent;
			const tab = app.setting.pluginTabs.find((t) => t.id === id);
			const defs = tab.getSettingDefinitions();
			const hasBuy = [...(app.setting.modalEl ?? document.body).querySelectorAll("button")].some((b) => b.textContent === "Buy Pro (19 USD)");
			app.setting.close();
			return {
				hasRows: ["Output folder", "Font size", "Page size", "License key", "Word template", "Page numbers"].every((s) => text.includes(s)),
				hasPro: text.includes("Pro"),
				hasBuy,
				defs: defs.length,
				items: defs.reduce((n, g) => n + (g.items?.length ?? 0), 0),
			};
		}, PLUGIN_ID);
		check("Settings tab renders (declarative API)", settings.hasRows && settings.hasPro && settings.hasBuy && settings.items === 10, JSON.stringify(settings));

		// Setting change through setControlValue persists to data.json
		await page.evaluate(async (id) => {
			const tab = app.setting.pluginTabs.find((t) => t.id === id);
			await tab.setControlValue("outputFolder", "  Exports  ");
		}, PLUGIN_ID);
		const data = JSON.parse(fs.readFileSync(path.join(VAULT, ".obsidian", "plugins", PLUGIN_ID, "data.json"), "utf8"));
		check("Setting saved and trimmed in data.json", data.outputFolder === "Exports", `outputFolder=${JSON.stringify(data.outputFolder)}`);
		const OUT = "Exports/Sample.docx";

		// 1. Command palette command
		let t0 = Date.now();
		const ran = await page.evaluate((id) => app.commands.executeCommandById(`${id}:export-current-note`), PLUGIN_ID);
		let file = await waitForFile(OUT, t0);
		check("Command 'Export current note to Word (.docx)' writes the file", ran && file, OUT);
		if (file) verifyDocx(file, "command");

		// 2. Ribbon icon
		fs.rmSync(path.join(VAULT, OUT));
		await sleep(500);
		t0 = Date.now();
		await page.click('.side-dock-ribbon-action[aria-label="Export note to Word"]');
		file = await waitForFile(OUT, t0);
		check("Ribbon icon writes the file", file, OUT);
		if (file) verifyDocx(file, "ribbon");

		// 3. File explorer right-click menu
		fs.rmSync(path.join(VAULT, OUT));
		await sleep(500);
		t0 = Date.now();
		await page.evaluate(() => app.workspace.leftSplit.expand());
		await page.evaluate(() => app.workspace.getLeavesOfType("file-explorer")[0] && app.workspace.revealLeaf(app.workspace.getLeavesOfType("file-explorer")[0]));
		await page.click('.nav-file-title[data-path="Sample.md"]', { button: "right" });
		await page.click('.menu-item:has-text("Export to Word (.docx)")');
		file = await waitForFile(OUT, t0);
		check("File menu item writes the file", file, OUT);
		if (file) verifyDocx(file, "file-menu");

		// Command is unavailable without a Markdown note
		const unavailable = await page.evaluate(async (id) => {
			await app.workspace.getLeaf(false).openFile(app.vault.getFileByPath("red.png"));
			await new Promise((r) => setTimeout(r, 300));
			// The command palette lists a command only when checkCallback(true) is true.
			const cmd = app.commands.findCommand(`${id}:export-current-note`);
			return app.workspace.getActiveFile()?.extension === "png" && cmd.checkCallback(true) === false;
		}, PLUGIN_ID);
		check("Command hidden when the active file is not a note", unavailable);

		// Pro features are gated while Pro is not active, and nothing calls the network
		const net = [];
		page.on("request", (r) => {
			if (!r.url().startsWith("app://") && !r.url().startsWith("data:") && !r.url().startsWith("blob:")) net.push(r.url());
		});
		await page.evaluate(async (id) => {
			const tab = app.setting.pluginTabs.find((t) => t.id === id);
			await tab.setControlValue("pageNumbers", true);
			await app.workspace.getLeaf(false).openFile(app.vault.getFileByPath("Sample.md"));
		}, PLUGIN_ID);
		await sleep(300);
		t0 = Date.now();
		await page.evaluate((id) => app.commands.executeCommandById(`${id}:export-current-note`), PLUGIN_ID);
		file = await waitForFile(OUT, t0);
		const gatedNotice = await page.evaluate(() => [...document.querySelectorAll(".notice")].some((n) => n.textContent.includes("Pro features")));
		const footer = file ? spawnSync("unzip", ["-p", file, "word/footer1.xml"], { encoding: "utf8" }).stdout : "";
		check("Page numbers stay off without Pro and a notice explains it", file && gatedNotice && !footer.includes("PAGE"));
		check("No network requests from the plugin", net.length === 0, net.join(" "));

		// Disable and compare with the state before enabling
		await waitNoticesGone(page);
		await page.evaluate((id) => app.plugins.disablePluginAndSave(id), PLUGIN_ID);
		await sleep(1000);
		await waitNoticesGone(page);
		// Remove the exported files so the file explorer matches the starting vault again.
		await page.evaluate(async () => {
			await app.workspace.getLeaf(false).openFile(app.vault.getFileByPath("Sample.md"));
			const out = app.vault.getFolderByPath("Exports");
			if (out) await app.vault.delete(out, true);
		});
		await sleep(800);
		const after = await snapshot(page, cdp);
		const diff = Object.keys(before).filter((k) => before[k] !== after[k]);
		check("Disable unloads the plugin", !after.loaded && after.commands === 0 && after.ribbonItems === 0 && after.settingTabs === 0);
		for (const k of ["workspaceHandlers", "vaultHandlers", "metadataHandlers", "listeners:window", "listeners:document", "listeners:document.body", "liveIntervals", "pluginNodes"]) {
			check(`Nothing left behind: ${k}`, before[k] === after[k], `${before[k]} -> ${after[k]}`);
		}
		check("Nothing left behind: DOM nodes", Math.abs(before.domNodes - after.domNodes) <= 0, `${before.domNodes} -> ${after.domNodes}`);
		fs.writeFileSync(path.join(RESULTS, "snapshots.json"), JSON.stringify({ before, enabled, after, diff }, null, 2));

		const pluginErrors = consoleErrors.filter((e) => /docx-styler|docx styler/i.test(e));
		check("No console errors or warnings from the plugin", pluginErrors.length === 0, pluginErrors.join(" | "));
		fs.writeFileSync(path.join(RESULTS, "console.txt"), consoleErrors.join("\n"));
	} finally {
		await browser?.close().catch(() => {});
		try {
			process.kill(-child.pid, "SIGTERM");
		} catch {
			/* already gone */
		}
		await sleep(1500);
		try {
			process.kill(-child.pid, "SIGKILL");
		} catch {
			/* already gone */
		}
	}
	const summary = { date: new Date().toISOString(), obsidian: OBSIDIAN_VERSION, passed: results.length - failed, failed, total: results.length, results };
	fs.writeFileSync(path.join(RESULTS, "summary.json"), JSON.stringify(summary, null, 2));
	console.log(`\nE2E: ${summary.passed}/${summary.total} passed. Results in ${RESULTS}`);
	process.exit(failed ? 1 : 0);
}

main().catch((e) => {
	console.error(e);
	process.exit(2);
});
