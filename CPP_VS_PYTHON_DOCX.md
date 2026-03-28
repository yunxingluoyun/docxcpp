# docxcpp vs python-docx

This file is a fresh comparison of the current C++ port in `cpp/` against the upstream `python-docx` implementation in this repository.

It is based on the code that exists today:

- C++ side: `cpp/include/docxcpp/*.hpp`, `cpp/src/*.cpp`, `cpp/tests/*.cpp`
- Python side: public behavior and feature surface reflected in `src/docx/`, `docs/user/`, and tests

Status labels used below:

- `Supported`: available in the current C++ port and covered by current tests
- `Partial`: available in a narrower or lower-level form than `python-docx`
- `Missing`: not currently exposed by the C++ port

## Current Position

`docxcpp` is no longer just a minimal exporter. It now covers a practical subset of Word document generation and inspection:

- open, create, and save `.docx`
- add paragraphs, headings, page breaks, and mixed-run content
- append runs to bound paragraph objects with `Paragraph::add_run()`
- write a useful subset of run formatting
- create tables, write cell content, append extra cell paragraphs, and merge cells
- insert inline pictures from file paths or byte buffers
- insert external hyperlinks
- write basic paragraph-anchored comments
- set page size, orientation, margins, header/footer distance, and gutter
- set paragraph alignment, indentation, spacing, fixed line spacing, and basic pagination flags
- read back paragraphs, runs, heading/style ids, paragraph formats, tables, comments, and picture metadata

At the same time, `python-docx` is still much broader. It offers a richer document object model, more editing-oriented APIs, deeper style/section coverage, and more complete support for Word concepts.

## Feature Snapshot

### Core document lifecycle

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Open existing document | Supported | Supported | Both can load `.docx`. |
| Create blank document | Supported | Supported | `docxcpp` uses the bundled template package. |
| Save document | Supported | Supported | Both write valid `.docx`. |

### Paragraphs and runs

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Add paragraph | Supported | Supported | `Document::add_paragraph()`. |
| Insert paragraph before another paragraph | Supported | Missing | No equivalent mutable paragraph API in C++. |
| Add heading/title | Supported | Supported | C++ supports heading levels `0-9`. |
| Add page break | Supported | Supported | `Document::add_page_break()`. |
| Mixed runs in one paragraph | Supported | Supported | C++ writes via `std::vector<Run>`. |
| Append run to an existing paragraph object | Supported | Partial | `Paragraph::add_run()` is available for document-bound paragraph handles, but the editable object model is still much smaller than `python-docx`. |
| Read paragraphs | Supported | Supported | `Document::paragraphs()`. |
| Read runs | Supported | Supported | `Paragraph::runs()`. |
| Read paragraph style info | Supported | Partial | C++ exposes `style_id()` and derived `heading_level()` only. |

### Text formatting

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Bold / italic / underline | Supported | Partial | C++ supports these three, but not the wider font feature set. |
| Font size | Supported | Supported | Integer points on the C++ side. |
| Font name | Supported | Supported | Basic `ascii` / `hAnsi` support. |
| Font color | Supported | Supported | RGB hex only in C++. |
| Highlight | Supported | Supported | Token-based Word highlight values. |
| Line breaks in run text | Supported | Supported | C++ writes inline line breaks and reads them back. |
| Tabs in run text | Supported | Partial | C++ reads `w:tab` as `\t`, but has no dedicated write-side tab-stop/text API. |
| Character styles by name | Supported | Missing | No style hierarchy API in C++. |

### Paragraph formatting

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Alignment | Supported | Partial | C++ supports left/center/right/justify/inherit. |
| Left/right/first-line indent | Supported | Supported | Read/write in points. |
| Space before / after | Supported | Supported | Read/write in points. |
| Line spacing | Supported | Partial | C++ currently supports fixed line spacing in points. |
| Pagination flags | Supported | Partial | C++ supports `keep_together`, `keep_with_next`, `page_break_before`. |
| Tab stops | Supported | Missing | Not exposed in current C++ API. |
| Widow/orphan control and richer paragraph flags | Supported | Missing | Not yet implemented. |
| Full paragraph format object model | Supported | Partial | `ParagraphFormat` exists, but is intentionally narrower. |

### Hyperlinks and comments

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| External hyperlink writing | Supported | Supported | Body and table-cell paragraph helpers exist. |
| Hyperlink object model on read | Supported | Partial | C++ exposes paragraph-level `HyperlinkInfo` snapshots via `Paragraph::hyperlinks()`, but not a full mutable hyperlink object graph. |
| Hyperlink target readback | Supported | Supported | Hyperlink URL readback is available through `Paragraph::hyperlinks()`. |
| Comment writing | Supported | Partial | C++ supports basic paragraph-anchored comment insertion. |
| Comment readback | Supported | Partial | `comments()` / `comment_count()` read the comments part, not full anchors/ranges. |
| Comment range selection on arbitrary runs | Supported | Missing | C++ does not support arbitrary run-range anchors. |
| Replies / resolved state / richer comment editing | Supported | Missing | Not implemented in C++. |

### Tables

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Add table | Supported | Supported | `Document::add_table()`. |
| Read tables | Supported | Supported | `Document::tables()`. |
| Write cell text | Supported | Supported | String and multi-run overloads are available. |
| Append extra paragraphs in a cell | Supported | Supported | `add_table_cell_paragraph()`. |
| Merge cells | Supported | Partial | Write support plus readback of `grid_span` / `vertical_merge`. |
| Table style selection | Supported | Partial | C++ writes fixed `TableGrid`; no public style selection API. |
| Add rows/columns after creation | Supported | Missing | No incremental structure editing. |
| Nested tables | Supported | Missing | Not supported in current C++. |
| Layout-grid helpers for omitted/merged cells | Supported | Missing | No `python-docx`-style helpers here. |

### Pictures

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Inline picture in body | Supported | Supported | PNG / JPEG currently. |
| Inline picture in table cell | Supported | Supported | Dedicated C++ helpers exist. |
| Picture sizing | Supported | Supported | Width/height in points, aspect ratio preserved when one side is omitted. |
| File-like image input | Supported | Partial | C++ uses byte-buffer APIs as the equivalent (`add_picture_data*`). |
| Floating shapes | Not supported | Not supported | Neither side exposes this as a normal high-level API. |
| Read picture metadata | Partial | Supported | C++ exposes size, relationship id, package target, and placement info. |
| Count images | Partial | Supported | `Document::image_count()`. |

### Sections and page layout

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Page size | Supported | Supported | Read/write in points. |
| Orientation | Supported | Supported | Portrait / landscape supported. |
| Margins | Supported | Supported | Top/right/bottom/left supported. |
| Header/footer distance | Supported | Supported | Included in `PageMargins`. |
| Gutter | Supported | Supported | Included in `PageMargins`. |
| Sections collection | Supported | Missing | C++ currently treats only the trailing section implicitly. |
| Add section / section break | Supported | Missing | No multi-section writer API. |
| Header/footer content | Supported | Missing | No header/footer parts or references yet. |

### Styles and document-wide subsystems

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Styles collection | Supported | Missing | No style catalog API in C++. |
| Custom style creation/deletion | Supported | Missing | Not implemented. |
| Apply named paragraph style | Supported | Partial | Heading styles only, indirectly. |
| Apply named table style | Supported | Missing | Fixed `TableGrid` only. |
| Settings object | Supported | Missing | No settings API in C++. |
| Broader document parts access | Supported | Missing | Coverage remains intentionally narrow. |

## Read-Side Reality

The read side is now meaningful, but still selective.

Current `docxcpp` read APIs cover:

- `Document::paragraphs()`
- `Paragraph::text()`
- `Paragraph::alignment()`
- `Paragraph::first_run_style()`
- `Paragraph::runs()`
- `Paragraph::style_id()`
- `Paragraph::heading_level()`
- `Paragraph::format()`
- `Paragraph::has_page_break()`
- `Paragraph::hyperlinks()`
- `Document::tables()`
- `TableRow::cells()`
- `TableCell::text()`
- `TableCell::grid_span()`
- `TableCell::vertical_merge()`
- `Document::page_size_pt()`
- `Document::page_orientation()`
- `Document::page_margins_pt()`
- `Document::comments()`
- `Document::comment_count()`
- `Document::picture_sizes_pt()`
- `Document::pictures()`
- `Document::image_count()`

What is still missing relative to `python-docx` is just as important:

- no section collection
- no header/footer read model
- no styles catalog
- no editable paragraph/run object graph
- no broader Word part coverage

## Biggest Gaps

The largest practical gaps compared with `python-docx` are:

1. Editing model
`python-docx` is built around mutable paragraph/run/table objects. `docxcpp` is still centered on document-level helper methods and lightweight snapshots.

2. Styles
`python-docx` has a real style system. `docxcpp` currently offers direct formatting plus heading-style shortcuts, but not a full style catalog.

3. Sections and headers/footers
`python-docx` supports sections as first-class objects. `docxcpp` currently supports only single-section page settings.

4. Hyperlink/comment depth
The C++ port now has useful basic support, but not the richer object model and anchor semantics available in `python-docx`.

5. Table richness
The current C++ table support is practical for generation, but still far from the fuller table DOM in `python-docx`.

## Recommended Next Steps

If the goal is to move the C++ port closer to `python-docx`, the highest-value next steps are:

1. Deepen paragraph editing further
- broaden the writable paragraph object model beyond `Paragraph::add_run()`
- append richer content types to existing paragraphs
- richer run/font properties

2. Deepen hyperlink/comment support
- richer hyperlink objects beyond `HyperlinkInfo`
- comment anchor range support beyond whole-paragraph insertion
- richer comment editing metadata

3. Expand paragraph formatting
- tab stops
- richer line-spacing modes
- widow/orphan control and more paragraph flags

4. Add section and header/footer support
- multiple sections
- section breaks
- header/footer references and content

5. Improve table APIs
- add row/column APIs
- nested tables
- better merged/omitted-cell helpers

6. Build a style system
- style catalog access
- named paragraph/table/character style application
- custom styles

## Bottom Line

`python-docx` remains the fuller and more editing-oriented library.

`docxcpp` is already useful for controlled document generation and inspection: text, mixed formatting, tables, pictures, page layout, basic hyperlinks, and basic comments are all in place. The remaining work is mainly about depth: richer editing objects, richer read models, styles, sections, and broader Word feature coverage.
