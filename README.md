# RichTextToHTML

A Google Apps Script that converts rich-text formatted cells in a Google Sheet into HTML.

## Setup

1. Open **Extensions → Apps Script** on the target spreadsheet and add `formatter.gs`.
2. On the active sheet, add two header columns: `RichText` and `HTMLText`.
3. Fill the `RichText` column with formatted text (bold, italic, strikethrough, underline, lists).

## Usage

Run `copyFormattedText` (from the Apps Script editor, or bind it to a menu/button). For each row:

* If `HTMLText` is already filled in, the row is skipped.
* Otherwise, the rich text in `RichText` is converted to HTML and written into `HTMLText`.

## Detected formatting and HTML output

* **Bold** → `<strong>...</strong>`
* _Italics_ → `<em>...</em>`
* Strikethrough → `<strike>...</strike>`
* _Underline_ → `<u>...</u>`
* Cell text is HTML-escaped (`&`, `<`, `>`) before any tags are applied, so literal angle-bracket text typed by a user is never rendered as raw HTML or mistaken for a formatting tag.
* No `<span>` elements are used — colour, font size, and hyperlinks are not detected or converted.
* Line breaks → raw newlines are converted to `<br />`, with cleanup to avoid unnecessary breaks immediately after closing list tags.

## List detection and conversion

After inline styling, the script collapses any formatting tag left wrapping only whitespace (an artifact of Google Sheets splitting a run right at the space after a list marker), then parses the text line-by-line to build ordered/unordered lists. It supports multiple marker styles and indent-based nested lists, ensuring proper opening/closing of `<ul>`/`<ol>` tags.

**Unordered list markers:**
* `- ` (dash-space) → wraps items in `<ul>...</ul>`.
* `* ` (asterisk-space) → also treated as unordered; rendered in `<ul>...</ul>`.

**Ordered list markers:**
* `N. ` (number-period-space) → `<ol>...</ol>`
* `a. ` (lowercase letter-period-space) → `<ol type="a">...</ol>`

A marker may be wrapped in a leading formatting tag (e.g. the whole line is bolded) — the script strips its own `<strong>`/`<em>`/`<strike>`/`<u>` tag before matching the marker, then reattaches it around the content.

**Nested lists:**
Indentation is measured in 2-space units; each 2-space increment is one nesting level, for any marker type (dash, star, number, or letter) — not just dash. A sublist always renders using its own marker's tag (e.g. a lettered sublist under a numbered parent still renders as `<ol type="a">`, a dashed sublist still renders as `<ul>`), regardless of the parent list's type.

**Automatic punctuation:**
If a list item's visible text doesn't end with a period, the script appends one before closing the `<li>` tag. It preserves any trailing HTML tags when inserting the period.

**Tag balancing and cleanup:**
Maintains a stack of open lists (tracking type, marker, and indent level) to emit correct closing tags when indentation decreases, when switching marker types at the same level, or when exiting list context entirely. After list processing, removes stray newlines directly after `</li>` `</ul>` `</ol>` and compresses newline-to-HTML-break conversions.

## Not supported

Text colour, font size, and hyperlinks are intentionally not detected or converted — only the formatting listed above is applied.
