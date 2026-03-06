# Word MCP Server

A fully-featured [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server for creating and manipulating Microsoft Word (`.docx`) files directly from Claude Code or any MCP-compatible AI assistant.

> **Author:** Renato Tadeu de Figueiredo
> **License:** MIT

---

## Features

| Category | Capabilities |
|---|---|
| **Document lifecycle** | Create, open, save, save-as, close, duplicate, auto-backup |
| **Reading** | Full content (paragraphs + tables), outline, section content, text search, headers/footers |
| **Text editing** | Replace text, replace paragraphs, replace/delete sections, insert/delete paragraphs, move sections |
| **Formatting** | Paragraph styles, alignment, spacing, indentation, line spacing, keep-together; character formatting (bold, italic, underline, strikethrough, font, size, color, all-caps, small-caps) |
| **Headers & Footers** | Set text, add page numbers (Page X / Page X of Y), insert logo images, clear, even/first-page variants |
| **Images** | Insert images in body or header, resize, align, add captions |
| **Tables** | Create, read, edit cells, add/delete rows, delete tables, merge cells (H/V), column widths, cell background & border formatting, header row |
| **Lists** | Bullet and numbered lists, multi-level |
| **Page layout** | Margins, orientation (portrait/landscape), page size (A4, A3, A5, Letter, Legal, custom), section breaks, page breaks, multi-column layout |
| **Styles** | List all styles, create/update custom paragraph styles |
| **Table of Contents** | Insert TOC field (auto-updated by Word) |
| **Hyperlinks** | Add clickable links to paragraphs |
| **Comments** | Add native Word comments (visible in review pane) |
| **Bookmarks** | Add named bookmarks |
| **Document properties** | Title, author, subject, keywords, description, category |
| **Advanced XML** | Direct OOXML patch for complex formatting |
| **Batch builder** | Create a complete document in one call from a JSON spec |

---

## Installation

### Prerequisites

- Python 3.11 or higher
- `pip` or `uv`

### 1. Clone the repository

```bash
git clone https://github.com/RenatoTadeuFigueiredo/word-mcp-server.git
cd word-mcp-server
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

Or with `uv`:

```bash
uv pip install -r requirements.txt
```

---

## Configuration

### Claude Code (CLI)

Add the server to your Claude Code MCP configuration. Run this command from the project folder:

```bash
claude mcp add word -- python /absolute/path/to/word-mcp-server/word_mcp_server.py
```

Or edit `~/.claude/mcp_servers.json` manually:

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["/absolute/path/to/word-mcp-server/word_mcp_server.py"]
    }
  }
}
```

Restart Claude Code after editing the configuration.

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or the equivalent path on Windows/Linux:

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["/absolute/path/to/word-mcp-server/word_mcp_server.py"]
    }
  }
}
```

---

## Usage Examples

The examples below show natural-language prompts you can give Claude once the MCP server is connected.

---

### 1. Create a complete company document in one call

```
Create a document at /Documents/company-report.docx with:
- Default font "Calibri", size 11
- Page margins: 2.5 cm all sides
- Title "Q1 2026 Business Report", author "Renato Figueiredo"
- Header: "Acme Corp — Confidential" (right-aligned)
- Footer: page numbers in "Page X of Y" format, centered
- Heading 1: "Executive Summary"
- A paragraph summarizing the quarter results
- Heading 1: "Financial Results"
- A table with columns: Quarter, Revenue, Expenses, Profit
  and three rows of data
- Heading 1: "Next Steps"
- A bullet list with three action items
```

Claude will use `build_document` to produce the full document in a single call.

---

### 2. Open and edit an existing document

```
Open /Documents/contract.docx, replace every occurrence of
"ACME Corp" with "Globex Corporation", and save.
```

```
Open /Documents/proposal.docx and change the font of all
paragraphs under the "Introduction" heading to
"Times New Roman" 12pt, justified alignment.
```

---

### 3. Add a company logo to the header

```
Open /Documents/report.docx, insert the image at
/images/logo.png into the header (width 4 cm, left-aligned),
and add a right-aligned footer with "Page X of Y".
```

---

### 4. Work with tables

```
Open /Documents/budget.docx and add a new row to the first table
with values: "Q4 2026", "$120,000", "$95,000", "$25,000".
```

```
In the open document, merge cells in table 0,
row 0, from column 0 to column 3 (full-width header row),
set its background color to "2E74B5" and text color to white.
```

---

### 5. Reorganize sections

```
In the open document, move the section "Appendix A"
to appear after the "Conclusion" section.
```

```
Delete the "Draft Notes" section entirely from the document.
```

---

### 6. Insert a Table of Contents

```
Open /Documents/manual.docx, insert a Table of Contents
at the beginning (after the title page) with the heading
"Contents", including up to level 3 headings.
```

> After opening the file in Word, press **Ctrl+A** then **F9** to render the page numbers.

---

### 7. Add comments for review

```
Open /Documents/draft.docx and add a comment from author
"Renato" to the paragraph containing "payment terms" with
the text: "Please confirm this clause with legal team."
```

---

### 8. Page layout adjustments

```
Open /Documents/brochure.docx, set the page to A4 landscape,
set 2 columns with 1.5 cm spacing between them.
```

```
Set page margins of the open document to:
top 3 cm, bottom 2.5 cm, left 2.5 cm, right 2 cm.
```

---

### 9. Create a custom style and apply it

```
In the open document, create a paragraph style called
"Highlight Box" based on "Normal", with bold text,
font size 11, color "2E74B5", and 6 pt spacing before and after.
Then apply it to all paragraphs containing the word "IMPORTANT".
```

---

### 10. Add hyperlinks

```
In the open document, find the paragraph containing
"visit our website" and add a hyperlink with text
"our website" pointing to https://example.com.
```

---

## Available Tools (41 total)

### Document Lifecycle
| Tool | Description |
|---|---|
| `open_document` | Open an existing `.docx` file |
| `create_new_document` | Create a new blank document (optional template) |
| `save_document` | Save (creates `.bak` backup by default) |
| `save_as` | Save to a new path |
| `close_document` | Close without saving |
| `duplicate_document` | Copy to a new path |

### Reading & Inspection
| Tool | Description |
|---|---|
| `get_document_info` | Metadata, word/char count, section count |
| `read_document` | Full structured JSON (paragraphs + tables) |
| `get_outline` | Heading hierarchy |
| `get_section` | Content under a specific heading |
| `find_text` | Search paragraphs and table cells |
| `read_table` | Read one or all tables |
| `list_styles` | List available styles |
| `get_document_xml` | Raw OOXML of paragraph or body |
| `read_headers_footers` | Read all headers/footers |

### Text Editing
| Tool | Description |
|---|---|
| `replace_text` | Find & replace (preserves formatting, searches tables too) |
| `replace_paragraph` | Replace a full paragraph's text |
| `replace_section` | Replace all content under a heading |
| `insert_paragraph` | Insert at position (after_text, after_heading, index, end) |
| `delete_paragraph` | Delete by index or text match |
| `delete_section` | Delete heading + all its content |
| `move_section` | Move a section before/after another |
| `add_heading` | Add a heading at any level |

### Formatting
| Tool | Description |
|---|---|
| `format_paragraph` | Style, alignment, spacing, indentation, line spacing |
| `format_text_run` | Bold, italic, underline, font, color, caps, strikethrough |
| `create_style` | Create/update a custom paragraph style |

### Headers & Footers
| Tool | Description |
|---|---|
| `set_header` | Set header text with formatting |
| `set_footer` | Set footer text + optional page numbers |
| `add_image_to_header` | Insert logo/image into header |
| `clear_header` | Remove header content |
| `clear_footer` | Remove footer content |
| `read_headers_footers` | Read header/footer content |

### Images
| Tool | Description |
|---|---|
| `insert_image` | Insert image in body with optional caption |

### Tables
| Tool | Description |
|---|---|
| `insert_table` | Create table with data and styling |
| `read_table` | Read table content |
| `edit_table_cell` | Edit cell text and formatting |
| `add_table_row` | Append a row |
| `delete_table_row` | Delete a row |
| `delete_table` | Delete entire table |
| `merge_table_cells` | Merge cells horizontally or vertically |
| `format_table_cell` | Set cell background and borders |

### Page Layout
| Tool | Description |
|---|---|
| `set_page_margins` | Set margins in cm |
| `set_page_orientation` | Portrait or landscape |
| `set_page_size` | A4, A3, A5, Letter, Legal, or custom |
| `add_page_break` | Insert hard page break |
| `add_section_break` | New page, continuous, even/odd page |
| `set_columns` | Multi-column layout |

### Advanced
| Tool | Description |
|---|---|
| `insert_toc` | Insert Table of Contents field |
| `add_hyperlink` | Add clickable link to a paragraph |
| `add_comment` | Add native Word comment |
| `add_bookmark` | Add named bookmark |
| `set_document_properties` | Title, author, subject, keywords, etc. |
| `apply_xml_patch` | Direct OOXML replacement |
| `build_document` | Create full document from JSON spec |

---

## Auto-backup

Every `save_document` call creates a `.bak` file next to the original by default. To disable:

```
Save the document without creating a backup.
```

---

## Notes

### Table of Contents
TOC fields are inserted as Word field codes. Open the file in Microsoft Word or LibreOffice Writer and press **Ctrl+A → F9** to render the actual page numbers.

### Comments
Native Word comments are injected into the `word/comments.xml` part. They appear in the Word review pane. If the comments part is inaccessible, a plain inline text fallback is used.

### Page Numbers
`PAGE` and `NUMPAGES` fields render correctly in Word and LibreOffice Writer.

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `python-docx` | ≥ 1.1.2 | Core `.docx` manipulation |
| `lxml` | ≥ 5.2.0 | XML processing |
| `mcp` | ≥ 1.0.0 | MCP server protocol |

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

## License

[MIT](LICENSE) — Copyright (c) 2026 Renato Tadeu de Figueiredo
