// esbuild plugin that keeps dynamic code and runtime script elements out of main.js.
//
// 1. jszip is bundled from its CommonJS source (lib/index.js) instead of dist/jszip.min.js,
//    with its polyfills (setimmediate, lie, readable-stream) swapped for small local shims.
// 2. docx ships a prebuilt dist/index.mjs that inlines its own copy of jszip.min.js and a
//    browser stream polyfill (readable-stream, util, get-intrinsic, function-bind). We point
//    docx at the cleaned jszip above and drop the stream polyfill, which only backs the unused
//    Packer.toStream API. Each edit must match exactly once, so a docx upgrade that changes
//    these lines fails the build instead of silently shipping the old code.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const shim = (name) => path.join(here, "shims", name);

const docxEdits = [
	{
		find: "var import_stream_browserify = require_stream_browserify();",
		replace:
			"var import_stream_browserify = { Stream: class { constructor() { throw new Error(\"Packer.toStream is not available in this build\"); } } };",
	},
	{
		find: "var import_jszip_min = /* @__PURE__ */ __toESM(require_jszip_min(), 1);",
		replace: "var import_jszip_min = { default: __cleanJSZip };",
	},
];

export const setImmediateShim = shim("set-immediate.js");

export const cleanDeps = {
	name: "clean-deps",
	setup(build) {
		const root = build.initialOptions.absWorkingDir ?? process.cwd();
		build.onResolve({ filter: /^jszip$/ }, () => ({
			path: path.join(root, "node_modules", "jszip", "lib", "index.js"),
		}));
		build.onResolve({ filter: /^setimmediate$/ }, () => ({ path: shim("empty.js") }));
		build.onResolve({ filter: /^(lie|immediate)$/ }, () => ({ path: shim("lie.js") }));
		build.onResolve({ filter: /^readable-stream$/ }, () => ({ path: shim("no-node-stream.js") }));

		build.onLoad({ filter: /[\\/]node_modules[\\/]docx[\\/]dist[\\/]index\.mjs$/ }, async (args) => {
			let code = await fs.promises.readFile(args.path, "utf8");
			for (const { find, replace } of docxEdits) {
				const count = code.split(find).length - 1;
				if (count !== 1) {
					throw new Error(`clean-deps: expected 1 match in docx for "${find}", found ${count}`);
				}
				code = code.replace(find, () => replace);
			}
			code = `import __cleanJSZip from "jszip";\n${code}`;
			return { contents: code, loader: "js" };
		});
	},
};
