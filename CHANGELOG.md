# Changelog

All notable changes to OneNoteXML will be documented in this file.

## [1.0.1] - 2025-11-20

### Added
- Nested page hierarchy support for Obsidian vaults
  - Pages now export to nested folder structure matching OneNote hierarchy
  - Parent pages get subfolders containing their child pages
  - Preserves OneNote's page levels (level 1, 2, 3, etc.)
- Page order preservation during export
  - XML files now prefixed with index numbers (001_, 002_, etc.)
  - Ensures pages process in correct hierarchical order, not alphabetically

### Fixed
- Leading underscores in section names now preserved (e.g., `_pure_test` stays `_pure_test`)
- Nested pages no longer flattened into single directory
- Page hierarchy correctly inferred from `pageLevel` attribute and export order

### Changed
- Export script (`export_xml_notebook.ps1`) now numbers pages sequentially
- Pipeline sorts pages by numeric prefix instead of alphabetically
- Obsidian converter builds parent-child relationships during conversion

## [1.0.0] - 2025-11-18

### Added
- Initial public release
- Direct XML extraction from OneNote via COM API
- Obsidian vault output with YAML frontmatter and wikilinks
- Logseq graph output with properties and queries
- CallbackID-based image extraction (64% success rate)
- Unified CLI entry point (`onenotexml.py`)
- Minimal dependencies (BeautifulSoup + pywin32)
- Production tested on 195 pages with 227 images

### Features
- Windows-only (OneNote COM API dependency)
- OneNote 2010-2013 support
- Local-only extraction (no cloud sync)
- Dual format output (Obsidian OR Logseq)

### Known Limitations
- OneNote has rich formatting, tables, columns, not all of that can be mapped nicely into Markdown format
- Hand-drawn ink not supported
- Windows-only platform support
- Requires OneNote 2010-2013 desktop version, but may work with other versions, although untested
