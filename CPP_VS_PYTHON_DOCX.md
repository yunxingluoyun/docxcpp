# docxcpp vs python-docx

This document compares the current C++ port under `cpp/` with the upstream `python-docx` feature set in this repository.

Comparison baseline:

- `docxcpp`: current implementation in `cpp/include/docxcpp/*.hpp`, `cpp/src/*.cpp`, and `cpp/tests/*.cpp`
- `python-docx`: user documentation under `docs/user/` and API/docs coverage in `src/docx/` and `tests/`

Status legend:

- `Supported`: implemented and covered by current C++ tests
- `Partial`: available in a narrower or lower-level form than `python-docx`
- `Not supported`: not currently exposed by the C++ port

## Summary

`docxcpp` already covers a practical subset of `python-docx`:

- create/open/save `.docx`
- add paragraphs, headings, page breaks
- set a limited set of run styles
- create tables, write cells, append cell paragraphs, merge cells
- insert inline pictures in body and table cells
- set page size and page margins
- read back paragraphs, runs, headings, tables, pictures, and basic section page settings

The current gap is that `docxcpp` is still a focused document generator/reader, while `python-docx` is a broader document object model with richer editing, styling, section, header/footer, and table APIs.

## Feature Matrix

| Area | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Open document | Supported | Supported | Both can open existing `.docx` files. |
| Create blank document from template | Supported | Supported | `docxcpp` uses the bundled default template directory. |
| Save document | Supported | Supported | Both can write `.docx`. |
| Append paragraph | Supported | Supported | `docxcpp::Document::add_paragraph()`. |
| Insert paragraph before existing paragraph | Supported | Not supported | `python-docx` has `Paragraph.insert_paragraph_before()`. |
| Append heading/title | Supported | Supported | `docxcpp` supports levels `0-9`. |
| Apply paragraph style by style name/object | Supported | Partial | `docxcpp` supports heading style ids indirectly and paragraph alignment, but not arbitrary named paragraph styles. |
| Add page break | Supported | Supported | `docxcpp::Document::add_page_break()`. |
| Read paragraph list | Supported | Supported | `docxcpp::Document::paragraphs()`. |
| Paragraph text | Supported | Supported | `Paragraph::text()`. |
| Paragraph alignment | Supported | Partial | `docxcpp` supports read/write for left/center/right/justify/inherit only. |
| Rich paragraph format object | Supported | Not supported | No C++ equivalent of `ParagraphFormat`. |
| Paragraph indentation | Supported | Not supported | No left/right/first-line indent API yet. |
| Paragraph spacing | Supported | Not supported | No before/after spacing API yet. |
| Line spacing | Supported | Not supported | No line spacing API yet. |
| Pagination flags (`keep_together`, etc.) | Supported | Not supported | No equivalent in C++ yet. |
| Paragraph style readback | Supported | Partial | `Paragraph::style_id()` and `Paragraph::heading_level()` only. |
| Run list | Supported | Supported | `Paragraph::runs()`. |
| Add run to existing paragraph | Supported | Not supported | `python-docx` exposes mutable `Paragraph.add_run()`. |
| Run text | Supported | Supported | Read side only in C++; writing always builds a single run per added paragraph. |
| Run bold/italic/underline | Supported | Partial | C++ supports these three, but not the broader font model. |
| Run font size | Supported | Supported | Current C++ uses point integers. |
| Run font name | Supported | Supported | Basic `ascii`/`hAnsi` write/read. |
| Run font color | Supported | Supported | RGB hex only. |
| Run highlight | Supported | Supported | Token-based highlight value support. |
| Tabs in run text | Supported | Partial | Read side maps `w:tab` to `\t`; no dedicated write-side tab-stop API. |
| Hyperlinks | Supported | Not supported | No hyperlink model or writer in C++. |
| Comments | Supported | Not supported | No comments part, anchors, metadata, or replies in C++. |
| Tables: add table | Supported | Supported | `docxcpp::Document::add_table()`. |
| Tables: style assignment | Supported | Partial | C++ writes a fixed `TableGrid` style, no public table-style API. |
| Tables: read tables | Supported | Supported | `docxcpp::Document::tables()`. |
| Tables: row/column collections | Supported | Partial | C++ exposes rows and cells, but no explicit column abstraction. |
| Tables: add row/column after creation | Supported | Not supported | No incremental structural editing API yet. |
| Tables: access cell by object model | Supported | Partial | C++ uses row/cell vectors on read and index-based setters on write. |
| Tables: cell text | Supported | Supported | `TableCell::text()` and `set_table_cell()`. |
| Tables: multiple paragraphs in cell | Supported | Supported | `add_table_cell_paragraph()` and readback concatenation. |
| Tables: nested tables | Supported | Not supported | No nested table writer/reader in C++. |
| Tables: merged cells | Supported | Partial | C++ supports merge writing and reads `grid_span`/`vertical_merge`, but lacks normalized layout-grid helpers. |
| Tables: omitted cells / layout-grid helpers | Supported | Not supported | No equivalent to `grid_cols_before` / `grid_cols_after`. |
| Inline pictures in body | Supported | Supported | JPEG/PNG only in current C++. |
| Inline pictures in table cell | Supported | Supported | Explicit C++ helper exists. |
| Picture sizing | Supported | Supported | Width/height in points, aspect ratio preserved when one side omitted. |
| File-like image input | Supported | Not supported | C++ currently accepts filesystem paths only. |
| Floating shapes | Not supported | Not supported | Neither side supports this in normal API. |
| Read picture metadata | Partial | Supported | C++ exposes size, rel id, package target, and placement info. |
| Count images | Partial | Supported | `docxcpp::Document::image_count()`. |
| Sections collection | Supported | Not supported | C++ currently treats only the trailing document section implicitly. |
| Add section | Supported | Not supported | No multi-section editing API in C++. |
| Page size | Supported | Supported | C++ reads/writes width and height. |
| Page orientation | Supported | Not supported | No orientation API in C++ yet. |
| Margins | Supported | Partial | C++ supports top/right/bottom/left only; no gutter/header/footer distances. |
| Header/footer content | Supported | Not supported | No header/footer parts or references in C++. |
| Styles collection | Supported | Not supported | No document style catalog API in C++. |
| Add/delete custom styles | Supported | Not supported | No equivalent in C++. |
| Apply table style by name | Supported | Not supported | Only fixed `TableGrid` output at the moment. |
| Apply character style by name | Supported | Not supported | No style hierarchy model in C++. |
| Document settings object | Supported | Not supported | No settings API in C++. |

## What docxcpp Does Well Already

Compared with many "first pass" ports, the current C++ implementation is already beyond simple text export:

- It produces valid `.docx` files with paragraphs, headings, page breaks, tables, and inline pictures.
- It supports useful direct formatting on generated text: bold, italic, underline, size, RGB color, font, highlight, and alignment.
- It can place pictures both in the document body and inside table cells.
- It can read back structured content rather than only writing:
  - paragraphs
  - runs
  - heading style ids
  - tables
  - merge metadata
  - picture metadata and placement
  - page size and margins
- It is already validated by C++ tests and by opening generated output through `python-docx`.

## Main Gaps Relative to python-docx

The largest missing areas are:

1. Rich text editing object model

- `python-docx` lets you mutate paragraph/run objects directly, append runs, set richer font properties, and work with hyperlinks.
- `docxcpp` is still centered on higher-level document methods and lightweight read models.

2. Full paragraph formatting

- `python-docx` exposes indentation, tab stops, paragraph spacing, line spacing, and pagination flags.
- `docxcpp` currently exposes only alignment and page-break detection.

3. Rich table model

- `python-docx` has a more complete table DOM, including row/column objects, incremental row insertion, nested tables, and helpers for non-uniform tables.
- `docxcpp` supports practical table creation and merge metadata, but not a full layout-grid abstraction.

4. Section and page-layout model

- `python-docx` supports multiple sections, section breaks, orientation, header/footer distances, and header/footer content access.
- `docxcpp` currently supports only single-section page size and four margins.

5. Styles and advanced document features

- `python-docx` has style catalogs, custom styles, comments, settings access, and broader document-part coverage.
- `docxcpp` does not yet expose those subsystems.

## Read-Side Comparison

On read APIs, the current situation is:

### python-docx

- broader and more object-oriented
- preserves more of the Word document model
- supports styles, sections, comments, headers/footers, and more detailed formatting APIs

### docxcpp

- strong on the subset it targets
- convenient for inspection of generated documents
- currently exposes these read models:
  - `Document::paragraphs()`
  - `Paragraph::text()`
  - `Paragraph::alignment()`
  - `Paragraph::first_run_style()`
  - `Paragraph::runs()`
  - `Paragraph::style_id()`
  - `Paragraph::heading_level()`
  - `Paragraph::has_page_break()`
  - `Document::tables()`
  - `TableRow::cells()`
  - `TableCell::text()`
  - `TableCell::grid_span()`
  - `TableCell::vertical_merge()`
  - `Document::page_size_pt()`
  - `Document::page_margins_pt()`
  - `Document::picture_sizes_pt()`
  - `Document::pictures()`
  - `Document::image_count()`

## Compatibility Notes

- `python-docx` is the reference implementation here; `docxcpp` is not yet API-compatible with it.
- The current C++ API is intentionally simpler and more index-based.
- Generated `.docx` files from the C++ side have been validated both by C++ tests and by reading them with `python-docx`.
- Table merge readback semantics differ from the higher-level approximations `python-docx` provides for merged cells.

## Suggested Next Priorities

If the goal is to move `docxcpp` closer to `python-docx`, the highest-value next steps are:

1. Add a mutable paragraph/run writer API
- `Paragraph::add_run()`
- mixed-format paragraphs
- richer run/font options

2. Expand paragraph formatting
- indent
- spacing
- line spacing
- page-break-before and related flags

3. Expand section APIs
- orientation
- multi-section support
- header/footer references and content

4. Improve table reading/writing
- add row
- nested tables
- normalized layout-grid helpers for merged/omitted cells

5. Add style system support
- document styles collection
- apply named paragraph/table/character styles
- custom style creation

## Bottom Line

`python-docx` is still much broader and closer to a full Word document object model.

`docxcpp` has already reached a useful milestone: it can generate and read back a meaningful subset of Word documents with text, formatting, tables, pictures, and basic page layout. For report generation, structured export, and controlled document templates, it is already practical. For full-feature parity with `python-docx`, the biggest missing work is in styles, sections, headers/footers, richer text editing, and advanced table handling.
