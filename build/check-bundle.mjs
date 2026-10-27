// Fails when main.js contains patterns the Obsidian community review flags:
// runtime script elements and dynamic code execution.
import fs from "node:fs";

const file = process.argv[2] ?? "main.js";
const code = fs.readFileSync(file, "utf8");

const banned = [
	{ name: "createElement(\"script\")", re: /createElement\(\s*["'`]script["'`]\s*\)/g },
	{ name: "eval(", re: /(^|[^\w$.])eval\s*\(/g },
	{ name: "new Function", re: /new\s+Function\b/g },
	{ name: "Function(...) call", re: /(^|[^\w$.])Function\s*\(/g },
	{ name: "setTimeout or setInterval with a string", re: /set(Timeout|Interval)\(\s*["'`]/g },
];

let failed = false;
for (const { name, re } of banned) {
	for (const match of code.matchAll(re)) {
		failed = true;
		const at = match.index ?? 0;
		const context = code.slice(Math.max(0, at - 60), at + 80).replace(/\s+/g, " ");
		console.error(`check:bundle: ${name} at offset ${at}: ...${context}...`);
	}
}

if (failed) {
	console.error(`check:bundle: ${file} failed`);
	process.exit(1);
}
console.log(`check:bundle: ${file} is clean (${code.length} bytes)`);
