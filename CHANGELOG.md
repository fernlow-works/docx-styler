# Changelog

## 1.3.0

- Pro: optional cover page for single notes and compiled folders. Title, subtitle and author line in Word's Title and Subtitle styles, then a page break. `title`, `subtitle` and `author` properties override the text. A compiled folder reads them from its folder note.
- Pro: optional Word table of contents built from the heading styles, levels 1 to 3 by default (1 to 6 possible), with a title in Word's TOC Heading style. It lists the headings before Word updates it, and Word refreshes it when the file opens.
- Both are off by default and can be turned on or off for one note with the `docx-cover` and `docx-toc` properties.
- Single note export without Pro is unchanged.

## 1.2.0

- Pro: compile a folder of notes into one Word file.
