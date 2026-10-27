# Docx Styler

Export Obsidian notes to Word (.docx) files that use real Word styles. Works on desktop and mobile. No Pandoc or other software to install.

## What you get for free

- Export the current note from the command palette, the ribbon, or a file's right-click menu.
- Headings become Word's Heading 1 to 6, so the Word navigation pane and table of contents work.
- Bold, italic, strikethrough, highlights, inline code and links.
- Bulleted, numbered, nested and task lists.
- Tables, block quotes, callouts and code blocks.
- Embedded vault images (`![[image.png|300]]` keeps the width).
- Footnotes become real Word footnotes.
- Wikilinks become their display text. Properties and `%%comments%%` are left out.
- A4 or Letter pages, your choice of font and size, files saved into your vault.

The free features never use the network and never stop working.

## Pro

Pro adds:

- **Your own Word template.** Point Docx Styler at a .docx in your vault and every export uses its styles (fonts, colors, spacing, heading looks).
- **Page numbers** in the footer.

Price: 19 USD, one-time. Use your key on up to 5 devices.

[Buy Docx Styler Pro on Gumroad](https://fernlowworks.gumroad.com/l/docxstylerpro). Your license key is in the Gumroad receipt email.

After buying, paste your key into **Settings > Docx Styler > License key** and click **Activate**.

The license check is in this public code, so anyone could remove it. Buying a key is how you support the plugin and keep it maintained.

## Payment and network use

- **Payment is required for full access.** Pro features need a license key sold through Gumroad.
- **Network use.** The plugin contacts only Gumroad's license verify API (`api.gumroad.com`), and only when you click **Activate** or **Check license**, to check your key. No telemetry, no ads, and remote images are not downloaded.

## Installation

- **From Community plugins:** open **Settings > Community plugins > Browse**, search for "Docx Styler", then install and enable it.
- **Manually:** download `main.js`, `manifest.json` and `styles.css` from the [latest release](https://github.com/fernlow-works/docx-styler/releases/latest) and copy them into `<your vault>/.obsidian/plugins/docx-styler/`. Then reload Obsidian and enable Docx Styler under **Settings > Community plugins**.

## Support

Found a bug or have an idea? Open an issue on [GitHub](https://github.com/fernlow-works/docx-styler/issues). For license key problems, open an issue without posting your key.

## Known limits

- A Word template must already contain the styles you want to change (for example Heading 1). Word only saves a style into a file after it has been used or modified there.
- Footnote definitions must be on one line.
- Note embeds (`![[Other note]]`) are shown as a label, not inlined. SVG and WebP images are shown as a label.

## Development

```bash
npm install
npm run build   # type check and bundle to main.js
npm test        # unit tests (vitest)
npm run lint    # Obsidian's official eslint rules
```

## License

MIT. See LICENSE.
