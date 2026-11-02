# Contributing to Docx Styler

Thanks for helping. Bug reports, ideas and small pull requests are all welcome.

## Reporting a bug

Open an issue on [GitHub](https://github.com/fernlow-works/docx-styler/issues) and include:

- Your Obsidian version and operating system.
- The Docx Styler version (Settings, Community plugins).
- A short sample note that shows the problem, and what you expected the .docx to look like.

Never post a license key in an issue. For license key problems, describe what happened and leave the key out.

## Building from source

You need Node.js 20 or newer.

```bash
npm install
npm run build   # type check, bundle and bundle safety check
npm run lint    # eslint with the official Obsidian rules
npm test        # unit tests
```

`npm run build` writes `main.js` next to `manifest.json` and `styles.css`. Copy those three files into `<your vault>/.obsidian/plugins/docx-styler/` to try your build.

## Pull requests

- Keep each pull request to one change, and add or update a unit test when behavior changes.
- `npm run lint`, `npm run build` and `npm test` must pass.
- Follow the [Obsidian plugin guidelines](https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines): no `innerHTML`, no global `app`, clean up on unload, sentence case in the UI.
- No telemetry, ads, obfuscated code or new network calls. The only remote call is the optional Pro license check described in the README.

By contributing you agree that your contribution is licensed under the MIT license in [LICENSE](LICENSE).
