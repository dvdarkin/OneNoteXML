#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNoteXML - Notebook-Specific Obsidian Vault Generator
Extract content from OneNote XML exports to Obsidian-compatible vault
Uses PowerShell COM exported XML files as input with notebook-specific paths
"""

from pathlib import Path
import sys

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from pipeline_base import (
    setup_logging,
    parse_pipeline_args,
    group_pages_by_section,
    discover_xml_files,
    log_pipeline_start,
    log_conversion_summary
)
from extractors.onenote_xml_parser import OneNoteXMLParser
from converters.obsidian_converter import ObsidianConverter

def process_section(section_name: str, xml_files: list, converter: ObsidianConverter, logger) -> tuple:
    """Process all pages in a section."""
    logger.info(f"Processing section: {section_name}")
    
    # Parse all pages in the section
    pages_data = []
    success_count = 0
    
    for xml_file in xml_files:
        try:
            logger.info(f"  Parsing: {xml_file.name}")
            
            # Parse XML
            parser = OneNoteXMLParser()
            parsed_data = parser.parse_page_xml(xml_file)
            
            logger.info(f"    - Page: {parsed_data['page_name']}")
            logger.info(f"    - Content items: {len(parsed_data['content'])}")
            logger.info(f"    - Images: {len(parsed_data['images'])}")
            
            pages_data.append(parsed_data)
            success_count += 1
            
        except Exception as e:
            logger.error(f"Error processing {xml_file}: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    # Convert section to Obsidian format
    if pages_data:
        section_data = {
            'section_name': section_name,
            'pages': pages_data
        }
        
        vault_path = converter.convert_section(section_data)
        logger.info(f"  Section converted to Obsidian vault")
        
    return success_count, len(xml_files)

def main():
    """Main Obsidian vault generation process."""
    # Parse command line arguments
    notebook_name, output_base_dir = parse_pipeline_args('obsidian_pipeline.py')

    # Setup logging
    logger = setup_logging(output_base_dir.parent, 'OneNoteObsidian')

    # Log pipeline start
    log_pipeline_start(logger, "Notebook-Specific Obsidian Vault Generator",
                      notebook_name, output_base_dir)

    # Construct paths
    xml_input_dir = output_base_dir / 'XML' / f'{notebook_name}_XML'
    obsidian_output_dir = output_base_dir / 'obsidian_vault'

    # Create Obsidian output directory
    obsidian_output_dir.mkdir(parents=True, exist_ok=True)

    # Discover XML files (includes validation and error handling)
    xml_files = discover_xml_files(xml_input_dir, logger)
    
    # Group files by section
    sections = group_pages_by_section(xml_files)
    logger.info(f"Found {len(sections)} section(s): {list(sections.keys())}")
    
    # Create Obsidian converter
    vault_name = f"{notebook_name}-Vault"
    converter = ObsidianConverter(obsidian_output_dir, vault_name)
    
    # Process each section
    total_success = 0
    total_files = 0
    
    for section_name, section_files in sections.items():
        success, total = process_section(section_name, section_files, converter, logger)
        total_success += success
        total_files += total
    
    # Save image dictionary for PowerShell extraction
    if total_success > 0:
        # Always try to save image dictionary, even if empty
        dict_file = converter.save_image_dictionary(obsidian_output_dir / "image_extraction_map.json")
        logger.info(f"Image dictionary saved: {dict_file}")
        
        if converter.image_dictionary:
            logger.info(f"Found {len(converter.image_dictionary)} images for extraction")
        else:
            logger.info("No images found in processed pages")
    
    # Create Obsidian configuration suggestions
    create_obsidian_config_guide(converter.vault_root, logger, notebook_name)
    
    # Summary
    log_conversion_summary(logger, total_success, total_files, sections)

    if total_success > 0:
        print(f"\nSuccessfully processed {total_success} page(s) from '{notebook_name}' notebook")
        print(f"Obsidian vault created at: {converter.vault_root}")
        
        # Show vault structure
        print("\nObsidian Vault Structure:")
        show_vault_structure(converter.vault_root)
        
        # Show image extraction info
        if converter.image_dictionary:
            print(f"\n📷 Found {len(converter.image_dictionary)} images to extract")
            print("Next step: Run extract_images.bat to download actual images")
            print("Images will be saved to: attachments/ folder")
        else:
            print("\n📷 No images found in processed pages")
        
        # Instructions
        print(f"\n🔮 To open in Obsidian:")
        print(f"1. Open Obsidian")
        print(f"2. Click 'Open folder as vault'")
        print(f"3. Select: {converter.vault_root}")
        print(f"4. Enjoy your imported OneNote content!")
        
    else:
        print(f"\n✗ No files processed successfully from '{notebook_name}' notebook")
    
    return 0 if total_success == total_files else 1

def show_vault_structure(vault_root: Path, max_depth: int = 2):
    """Display the vault structure."""
    def show_tree(path: Path, prefix: str = "", depth: int = 0):
        if depth > max_depth:
            return
            
        items = sorted([
            item for item in path.iterdir() 
            if not item.name.startswith('.') and item.name != '__pycache__'
        ], key=lambda x: (x.is_file(), x.name.lower()))
        
        for i, item in enumerate(items):
            is_last = i == len(items) - 1
            current_prefix = "└── " if is_last else "├── "
            
            if item.is_dir():
                print(f"{prefix}{current_prefix}{item.name}/")
                next_prefix = prefix + ("    " if is_last else "│   ")
                show_tree(item, next_prefix, depth + 1)
            else:
                print(f"{prefix}{current_prefix}{item.name}")
    
    print(f"{vault_root.name}/")
    show_tree(vault_root)

def create_obsidian_config_guide(vault_root: Path, logger, notebook_name: str):
    """Create a configuration guide for Obsidian setup."""
    guide_content = f"""# Obsidian Configuration Guide - {notebook_name} Notebook

This vault was generated from OneNote '{notebook_name}' notebook export. Here are recommended settings:

## Recommended Obsidian Settings

### Files & Links
- **Default location for new attachments**: In subfolder under current folder
- **Subfolder name**: attachments
- **Use [[Wikilinks]]**: Enabled (recommended for full Obsidian features)
- **Automatically update internal links**: Enabled

### Editor
- **Spellcheck**: Enabled
- **Readable line length**: Enabled
- **Strict line breaks**: Disabled

### Core Plugins to Enable
- [ ] **Templates** - For creating new notes with consistent structure
- [ ] **Daily notes** - If you imported diary/journal content
- [ ] **Graph view** - Visualize connections between notes
- [ ] **Backlinks** - See what links to current note
- [ ] **Tag pane** - Browse notes by tags
- [ ] **Search** - Enhanced search capabilities

### Community Plugins (Optional)
- **Dataview** - Query your notes like a database
- **Calendar** - Calendar view for daily notes
- **Advanced Tables** - Better table editing
- **Image Auto Upload** - Automatic image handling

## Tags Used in Import
- `#onenote-import` - All imported notes from {notebook_name}
- Additional tags based on your OneNote structure

## Next Steps
1. Explore the **Graph view** to see connections
2. Use **Search** to find content across all notes
3. Create an **Index note** to organize your imported content
4. Set up **Templates** for new notes
5. Configure **Daily notes** if you have diary content

---
*Generated during OneNote '{notebook_name}' import*
"""
    
    guide_path = vault_root / f"00-{notebook_name}-Setup-Guide.md"
    with open(guide_path, 'w', encoding='utf-8') as f:
        f.write(guide_content)
    
    logger.info(f"Created Obsidian setup guide: {guide_path.name}")

if __name__ == "__main__":
    sys.exit(main())