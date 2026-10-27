// Stand-in for the setimmediate and immediate npm polyfills.
// Those packages fall back to createElement("script") and new Function(),
// which the Obsidian community review flags. This version only uses
// the native setImmediate (desktop), MessageChannel or setTimeout.
const queue = [];
let channel = null;

function flush() {
	const tasks = queue.splice(0, queue.length);
	for (const task of tasks) task();
}

export function setImmediate(callback, ...args) {
	const native = globalThis.setImmediate;
	if (typeof native === "function" && native !== setImmediate) {
		return native(callback, ...args);
	}
	queue.push(() => callback(...args));
	if (typeof MessageChannel === "function") {
		if (!channel) {
			channel = new MessageChannel();
			channel.port1.onmessage = flush;
		}
		channel.port2.postMessage(0);
	} else {
		setTimeout(flush, 0);
	}
}
