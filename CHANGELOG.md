# Changelog

All notable changes to OneNoteXML will be documented in this file.

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
