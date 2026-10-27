// Replaces the lie Promise polyfill used by jszip. Every Obsidian runtime has a native Promise,
// and lie pulls in the immediate package, which creates script elements at runtime.
module.exports = Promise;
