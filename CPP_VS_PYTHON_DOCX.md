# docxcpp vs python-docx

This document summarizes where the current C++ port in `cpp/` stands relative to upstream `python-docx`.

## Short Version

`docxcpp` is already beyond a toy exporter. It can now generate and read a practical subset of `.docx` features:

- paragraphs, headings, runs
- paragraph formatting
- tables and basic merges
- pictures
- hyperlinks
- basic comments
- page layout
- a meaningful read-side API

It is still not a full replacement for `python-docx`. The Python library remains much stronger in object model depth, styles, sections, headers/footers, and table richness.

## Status Labels

- `Supported`: implemented now and covered by current tests
- `Partial`: present, but clearly narrower than `python-docx`
- `Missing`: not exposed or not implemented

## What `docxcpp` Already Does Well

The current C++ code supports:

- open, create, and save `.docx`
- add paragraphs, headings, and page breaks
- write mixed-run paragraphs and headings
- append to an existing paragraph using `Paragraph::add_run()`
- basic run formatting: bold, italic, underline, size, font, color, highlight
- paragraph formatting: alignment, indentation, spacing, fixed line spacing, pagination flags
- create tables, write cells, append extra cell paragraphs, merge cells
- insert pictures from file paths and in-memory byte buffers
- add hyperlinks in body paragraphs and table cells
- add basic paragraph-level comments
- set page size, orientation, margins, header/footer distance, and gutter
- read paragraphs, runs, style ids, heading levels, paragraph formatting, tables, pictures, comments
- read `sections()`, including multiple sections on the read side

## Feature Comparison

### Document lifecycle

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Open document | Supported | Supported | Both load `.docx` |
| Create document | Supported | Supported | C++ uses a bundled template |
| Save document | Supported | Supported | Both produce valid `.docx` |

### Paragraphs and runs

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Add paragraph | Supported | Supported | `Document::add_paragraph()` |
| Add heading | Supported | Supported | Heading levels supported |
| Add page break | Supported | Supported | `Document::add_page_break()` |
| Multiple runs in one paragraph | Supported | Supported | `std::vector<Run>` |
| Append run to existing paragraph | Supported | Partial | `Paragraph::add_run()` exists, but the writable object model is still small |
| Insert paragraph before another paragraph | Supported | Missing | No equivalent C++ API yet |
| Read paragraphs | Supported | Supported | `Document::paragraphs()` |
| Read runs | Supported | Supported | `Paragraph::runs()` |

### Run/text formatting

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Bold/italic/underline | Supported | Supported | Basic formatting is present |
| Font size | Supported | Supported | Point-based |
| Font name | Supported | Supported | Basic font-name writing |
| Font color | Supported | Supported | RGB hex |
| Highlight | Supported | Supported | Word highlight tokens |
| Inline line breaks | Supported | Supported | Write and read supported |
| Tab stops | Supported | Missing | No API yet |
| Character styles | Supported | Missing | No style system yet |

### Paragraph formatting

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Alignment | Supported | Supported | left/center/right/justify |
| Left/right/first-line indent | Supported | Supported | Read/write available |
| Space before/after | Supported | Supported | Read/write available |
| Line spacing | Supported | Partial | Fixed line spacing only today |
| keep_together / keep_with_next / page_break_before | Supported | Supported | Implemented |
| Widow/orphan control | Supported | Missing | Not implemented |

### Tables

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Add table | Supported | Supported | `Document::add_table()` |
| Read tables | Supported | Supported | `Document::tables()` |
| Write cell text | Supported | Supported | Text and multi-run overloads |
| Append paragraphs in cell | Supported | Supported | Implemented |
| Merge cells | Supported | Partial | Basic merge write/read support |
| Add rows/columns after creation | Supported | Missing | Not available yet |
| Nested tables | Supported | Missing | Not available yet |

### Pictures

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Body pictures | Supported | Supported | PNG/JPEG |
| Cell pictures | Supported | Supported | Implemented |
| Picture sizing | Supported | Supported | Point-based |
| File-like input | Supported | Partial | C++ uses byte-buffer APIs as the equivalent |
| Read picture metadata | Partial | Supported | C++ exposes direct metadata helpers |

### Hyperlinks and comments

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Write hyperlinks | Supported | Supported | Body and table-cell helpers |
| Read hyperlink target | Supported | Supported | `Paragraph::hyperlinks()` |
| Hyperlink object model | Supported | Partial | Current C++ shape is a lightweight snapshot |
| Write comments | Supported | Partial | Basic paragraph-level insertion |
| Read comments | Supported | Partial | Comments-part metadata only |
| Arbitrary run-range comment anchors | Supported | Missing | Not implemented |

### Page layout and sections

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Page size | Supported | Supported | Read/write |
| Orientation | Supported | Supported | portrait/landscape |
| Margins | Supported | Supported | Read/write |
| Header/footer distance | Supported | Supported | Implemented |
| Gutter | Supported | Supported | Implemented |
| Read `sections()` collection | Supported | Partial | Multi-section readback now exists |
| Add section / section break | Supported | Missing | No writer API yet |
| Header/footer content | Supported | Missing | No object model yet |

### Styles

| Feature | python-docx | docxcpp | Notes |
| --- | --- | --- | --- |
| Styles collection | Supported | Missing | Not implemented |
| Apply named paragraph style | Supported | Partial | Mostly heading shortcuts today |
| Character/table styles | Supported | Missing | Not implemented |
| Custom styles | Supported | Missing | Not implemented |

## Read-Side Status

The read side is now real, not placeholder-only. Current C++ read APIs cover:

- paragraph text
- paragraph alignment
- runs
- style ids
- heading levels
- paragraph formatting
- page settings
- `sections()`
- tables and cells
- picture metadata
- comment metadata
- hyperlink text and targets

But it still lacks:

- a full styles system
- header/footer read models
- a richer editable paragraph/run object graph
- broader Word part coverage

## Highest-Value Gaps

If the goal is to move `docxcpp` closer to `python-docx`, the most important next areas are:

1. richer paragraph/run writing APIs
2. deeper hyperlink/comment object models
3. paragraph formatting expansion, especially tab stops and richer line spacing
4. section breaks, multi-section writing, and header/footer support
5. stronger table APIs
6. a real styles system

## Bottom Line

`python-docx` is still the deeper and more complete library.

`docxcpp` is already useful for real document automation workloads, especially structured generation, template filling, and controlled readback. The next big step is depth: richer object models, styles, sections, headers/footers, and broader editing coverage.
