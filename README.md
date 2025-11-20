# OneNoteXML

**Direct XML extraction from OneNote to Obsidian/Logseq markdown.**

Minimal dependencies, fully local, files never leave your computer. Just working image extraction and clean markdown.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

---

## Why Another OneNote Converter?

I tried existing tools to migrate my OneNote notebooks and got mixed results. 
So I built **OneNoteXML** - direct XML extraction using OneNote's COM API. XML is great to work with and allows for full flexibility: explore your data and decide which formatting and content you want to convert and how.


**If the other tools worked for you, great!** But if you're here because they didn't... this might help.

---

## What Makes It Different

- **Direct XML extraction** - No lossy Word→Pandoc conversion chain
- **Working image extraction** - CallbackID-based with fallbacks
- **Dual format support** - Output Obsidian vaults OR Logseq graphs
- **Minimal dependencies** - Just BeautifulSoup + pywin32
- **Local-only** - No cloud sync or Microsoft authentication required
- **Production tested** - Successfully converted 195 pages with 227 images

---

## Format Support Status

**Obsidian** (Primary platform - fully supported):
- Nested page hierarchy (folder structure)
- YAML frontmatter with metadata
- Wikilinks for internal references
- Image embedding with proper paths
- Actively maintained (I use Obsidian daily)

**Logseq** (Basic support):
- Page export with properties
- Block-based content structure
- Task detection (TODO/DONE)
- Flat page structure (no hierarchy)
- Limited testing (contributions welcome)

**Why the difference?** I built this tool to migrate my notebooks to Obsidian. That's what I use, test, and maintain. Logseq support exists because I couldn't choose initially what would work for me. Now that I have, I can't guarantee feature parity or catch platform-specific issues.

**If you need better Logseq support:** Pull requests welcome! The XML parser is format-agnostic, so improving the Logseq converter is straightforward if you know Logseq's conventions.

---

## Quick Start

### 1. Install

```bash
git clone https://github.com/dvdarkin/OneNoteXML.git
cd OneNoteXML
pip install -r requirements.txt
```

### 2. Check Requirements

```bash
python onenotexml.py --check-only
```

This verifies:
- Windows OS
- Python 3.8+
- OneNote 2010-2013 accessible via COM

### 3. Extract Notebook

```bash
# Extract to Obsidian format (default)
python onenotexml.py "Personal"

# Extract to Logseq format
python onenotexml.py "Work Notes" --format logseq

# Custom output directory
python onenotexml.py "Research" --output ./my-vault
```

### 4. Open in Obsidian/Logseq

**Obsidian:**
1. Open Obsidian
2. Click "Open folder as vault"
3. Select: `output/YourNotebook/YourNotebook-Vault`

**Logseq:**
1. Open Logseq
2. Add graph
3. Select: `output/YourNotebook/logseq_vault`

---

## Requirements

**Platform:**
- Windows (OneNote COM API dependency)
- OneNote 2010-2013 desktop version
- Python 3.8+

**Python packages:**
```bash
pip install -r requirements.txt
```

---

## How It Works

```
OneNote Notebook
    ↓
[1] PowerShell COM API → Raw XML files
    ↓
[2] Python XML Parser → Structured content
    ↓
[3] Format Converter → Obsidian OR Logseq markdown
    ↓
[4] Image Extractor → Binary image files via CallbackID
    ↓
Complete vault ready for use
```

**Technical approach:**
- Uses OneNote's COM Object Model to export raw XML
- Parses XML directly (no Word intermediate format)
- Extracts images via `GetBinaryPageContent(callbackID)`
- Converts to Obsidian (YAML frontmatter, wikilinks) or Logseq (properties, queries)

---

## Example Output

### Obsidian Format
```markdown
---
title: Project Notes
tags: [onenote-import, projects]
onenote_source: Work > Projects > Q4
created: 2024-11-15
---

# Project Notes

![[project-diagram.png]]

## Key Milestones
- [x] Phase 1 complete
- [ ] Phase 2 in progress

[[Related Note]] - See also: [[Timeline]]
```

### Logseq Format
```markdown
- # Project Notes
- Properties:
  notebook:: [[Work]]
  section:: [[Projects]]
  created:: [[2024-11-15]]
- ![project diagram](../assets/project-diagram.png)
- ## Key Milestones
  - DONE Phase 1 complete
  - TODO Phase 2 in progress
- [[Related Note]] - See also: [[Timeline]]
```


---

## Troubleshooting

### "Cannot access OneNote"
- Install OneNote 2010-2013 desktop version (not Windows 10 app)
- Open OneNote at least once to initialize COM registration
- Try running PowerShell/Python as Administrator

### "Notebook not found"
- Check spelling (case-sensitive)
- Open notebook in OneNote first
- Ensure notebook is downloaded locally (not web-only)
- Run with incorrect name to see list of available notebooks

### "Images not extracted"
- Supports JPG and PNG files that are embedded, other formats may fail
- Check `logs/` directory for details
- Some images stored externally by OneNote
- Original OneNote notebook remains unchanged (safe to retry)

---

## Contributing

This is a personal tool I built for myself. It works for my use cases. I will be improving it when I discover fidelity issues in conversions, but I'll mostly be doing it at my own pace.

**Bug reports welcome** - Open an issue with:
- Notebook structure (pages/sections)
- Error messages from console
- Log files from `logs/` directory

**Pull requests considered** - But please open an issue first to discuss

**Feature requests** - Probably won't implement, but feel free to fork

---

## License

MIT License - see [LICENSE](LICENSE) file.

Built because I needed to migrate my notebooks and existing tools failed. Open sourced in case it helps someone else.

---

## FAQ

**Q: Why not use Obsidian's official importer?**
A: Requires cloud sync. This tool works with local notebooks only.

**Q: Will you add Mac support?**
A: No. OneNote's COM API is Windows-only. Mac users: try running in Parallels/VM.

**Q: Can I export to other formats?**
A: Currently Obsidian and Logseq only. The XML parser is separate, so you could write converters for other formats.

**Q: Does this modify my OneNote notebooks?**
A: No. Read-only extraction. Your original notebooks are untouched.

**Q: What about newer OneNote formats?**
A: Not tested, you're welcome to test and report back.

**Q: Why does my notebook content look so bad?**
A: OneNote has very rich formatting capabilities not all of that can be mapped nicely into Markdown format. I only added support for what I've seen in my notebooks. You can run the tool with --debug flag to explore raw XML to see how it looks at the core if something is not working for you.

**Q: Why is Logseq support limited compared to Obsidian?**
A: I use Obsidian daily, so that's what gets tested and improved. Logseq support works for basic use cases, but I can't maintain features for a platform I don't use. If you need better Logseq support, contributions are welcome - the architecture makes it easy to improve the Logseq converter independently.

---

**Status:** Working tool. Use at your own risk. No warranty. MIT licensed.
