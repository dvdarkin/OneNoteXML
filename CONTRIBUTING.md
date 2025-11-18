# Contributing to OneNoteXML

Thank you for considering contributing to OneNoteXML! This document provides guidelines for contributing to the project.

## Table of Contents

- [Reporting Bugs](#reporting-bugs)
- [Suggesting Features](#suggesting-features)
- [Code Contributions](#code-contributions)
- [Testing](#testing)
- [Documentation](#documentation)
- [Questions](#questions)

---

## Reporting Bugs

Found a bug? Please [open an issue](https://github.com/dvdarkin/OneNoteXML/issues/new) and include:

### Essential Information

- **Windows version** (e.g., Windows 10, Windows 11)
- **OneNote version** (must be 2010-2013 desktop edition)
- **Python version** (`python --version`)
- **Installation method** (pip, git clone, etc.)

### Logs and Output

- **Full error logs** from `logs/` directory (redact personal info)
- **Console output** showing the complete error message
- **Notebook structure** (number of sections, pages, images)

### Reproducible Example

If possible, provide:
- Sample XML file (you can redact content, keep structure)
- Steps to reproduce the issue
- Expected vs. actual behavior

**Note:** Please check [existing issues](https://github.com/dvdarkin/OneNoteXML/issues) before creating a new one.

---

## Suggesting Features

We welcome feature requests! Please open an issue describing:

- **What you're trying to achieve** - Your use case
- **Why current functionality doesn't work** - What's missing
- **Proposed solution** - How you envision it working
- **Alternatives considered** - Other approaches you've thought about

### Feature Scope

OneNoteXML focuses on:
- Direct XML extraction from OneNote COM API
- Conversion to Obsidian and Logseq formats
- Image extraction via CallbackID mechanism
- Windows-based OneNote 2010-2013 desktop

Out of scope:
- OneNote cloud/web versions (no COM API)
- Cross-platform support (COM API is Windows-only)
- Real-time sync (this is a one-time migration tool)
- Other output formats (focus is Obsidian/Logseq)

---

## Code Contributions

### Getting Started

1. **Fork the repository**
   ```bash
   # On GitHub: Click "Fork" button
   ```

2. **Clone your fork**
   ```bash
   git clone https://github.com/YOUR_USERNAME/OneNoteXML.git
   cd OneNoteXML
   ```

3. **Create a virtual environment** (optional but recommended)
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

4. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

5. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

### Code Standards

#### Python Style

- Follow [PEP 8](https://pep8.org/) style guide
- Use descriptive variable names (`page_content` not `pc`)
- Add docstrings to all functions and classes
- Keep functions focused (single responsibility)

#### Example Function

```python
def extract_page_title(xml_root: ET.Element) -> str:
    """
    Extract page title from OneNote XML root element.

    Args:
        xml_root: Parsed XML root element

    Returns:
        Page title as string, or "Untitled" if not found
    """
    # Implementation...
```

#### PowerShell Style

- Use descriptive parameter names
- Add comment-based help at top of script
- Handle errors gracefully (don't crash on missing images)

### Commit Messages

Use clear, descriptive commit messages:

```bash
# Good
git commit -m "Fix image extraction for PNG files with transparency"
git commit -m "Add support for nested table structures"

# Less good
git commit -m "fix bug"
git commit -m "update"
```

### Pull Request Process

1. **Update documentation** if you changed functionality
2. **Test with real OneNote notebooks** (see Testing section)
3. **Update CHANGELOG.md** with your changes
4. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```
5. **Open a Pull Request** on GitHub
   - Describe what you changed and why
   - Reference any related issues (#123)
   - Include test results

---

## Testing

### Testing Requirements

Since OneNoteXML requires OneNote COM API, testing requires:
- Windows OS
- OneNote 2010-2013 desktop edition
- Python 3.8+

### Manual Testing Checklist

Before submitting a PR, test with:

- [ ] **Multiple notebook types**
  - Personal notebooks (diary, journal)
  - Work notebooks (structured, hierarchical)
  - Mixed content (text, images, tables, lists)

- [ ] **Both output formats**
  - Obsidian vault generation
  - Logseq graph generation

- [ ] **Image handling**
  - Notebooks with many images
  - Different image types (PNG, JPG, GIF)
  - Embedded vs. attached images

- [ ] **Edge cases**
  - Empty sections
  - Pages with no title
  - Special characters in filenames
  - Very long page names

### Running Tests

```bash
# Check requirements
python onenotexml.py --check-only

# Test extraction with a small notebook
python onenotexml.py "TestNotebook" --format obsidian

# Test Logseq format
python onenotexml.py "TestNotebook" --format logseq

# Check logs for errors
cat logs/*.log
```

### What Good Test Results Look Like

```
All requirements met

[1/3] Exporting XML from OneNote...
      XML export completed

[2/3] Converting to obsidian format...
      Obsidian conversion completed

[3/3] Extracting images...
      Image extraction completed
      Images copied: Y of X

[Verification]
      Created X markdown files
      Copied Y images to vault
```

---

## Documentation

### Documentation Types

1. **Code documentation** - Docstrings in Python files
2. **User documentation** - README.md, README_USAGE.md
3. **Research documentation** - Files in `docs/` directory

### Updating Documentation

If you change functionality:

- **README.md** - Update feature list, examples, or requirements
- **README_USAGE.md** - Update usage instructions
- **CHANGELOG.md** - Add entry for your change
- **Docstrings** - Keep function documentation accurate

### Documentation Style

- Use **bold** for UI elements and important concepts
- Use `code formatting` for commands and file paths
- Include **examples** showing real usage
- Keep it **concise** - users want quick answers

---

## Questions?

### Where to Ask

- **General questions** - Open a [GitHub Discussion](https://github.com/dvdarkin/OneNoteXML/discussions)
- **Bug reports** - Open an [Issue](https://github.com/dvdarkin/OneNoteXML/issues)
- **Feature ideas** - Open an [Issue](https://github.com/dvdarkin/OneNoteXML/issues) with "Feature Request" label

### Response Time

This is a personal project maintained in spare time. I'll do my best to respond within:
- **Critical bugs** - 1-3 days
- **Feature requests** - 1-2 weeks
- **General questions** - 1 week

### Getting Help

If you're stuck:
1. Check the [README.md troubleshooting section](README.md#troubleshooting)
2. Search [existing issues](https://github.com/dvdarkin/OneNoteXML/issues)
3. Look at `docs/` for technical deep-dives
4. Open a new issue with details

---

## Code of Conduct

### Our Pledge

We want OneNoteXML to be a welcoming project for everyone. We pledge to:
- Be respectful and inclusive
- Welcome newcomers and help them contribute
- Focus on what's best for the community
- Accept constructive criticism gracefully

### Unacceptable Behavior

- Personal attacks or insults
- Harassment of any kind
- Publishing others' private information
- Spam or off-topic comments

### Enforcement

If you experience or witness unacceptable behavior, please contact the project maintainer. Violations may result in temporary or permanent ban from the project.

---

## License

By contributing to OneNoteXML, you agree that your contributions will be licensed under the MIT License.

---

## Thank You!

Every contribution helps make OneNoteXML better for everyone trying to migrate from OneNote. Whether you:
- Report a bug
- Suggest a feature
- Fix a typo
- Write code

Your help is appreciated.
