#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNoteXML - Logseq Graph Generator
Extract content from OneNote XML exports to Logseq-compatible format
Uses PowerShell COM exported XML files as input
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
from converters.logseq_converter import LogseqConverter

def process_section(section_name: str, xml_files: list, converter: LogseqConverter, logger, notebook_name: str) -> tuple:
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
            
            # Add notebook name to parsed data
            parsed_data['notebook_name'] = notebook_name
            
            logger.info(f"    - Page: {parsed_data['page_name']}")
            logger.info(f"    - Content items: {len(parsed_data['content'])}")
            logger.info(f"    - Images: {len(parsed_data['images'])}")
            
            pages_data.append(parsed_data)
            success_count += 1
            
        except Exception as e:
            logger.error(f"Error processing {xml_file}: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    # Convert section to Logseq format
    if pages_data:
        section_data = {
            'section_name': section_name,
            'pages': pages_data
        }
        
        graph_path = converter.convert_section(section_data)
        logger.info(f"  Section converted to Logseq format")
        
    return success_count, len(xml_files)

def main():
    """Main Logseq graph generation process."""
    # Parse command line arguments
    notebook_name, output_base_dir = parse_pipeline_args('logseq_pipeline.py')

    # Setup logging
    logger = setup_logging(output_base_dir.parent, 'OneNoteLogseq')

    # Log pipeline start
    log_pipeline_start(logger, "Logseq Graph Generator",
                      notebook_name, output_base_dir)

    # Construct paths
    xml_input_dir = output_base_dir / 'XML' / f'{notebook_name}_XML'
    logseq_output_dir = output_base_dir / 'logseq_vault'

    # Create Logseq output directory
    logseq_output_dir.mkdir(exist_ok=True)

    # Discover XML files (includes validation and error handling)
    xml_files = discover_xml_files(xml_input_dir, logger)
    
    # Group files by section
    sections = group_pages_by_section(xml_files)
    logger.info(f"Found {len(sections)} section(s): {list(sections.keys())}")
    
    # Create Logseq converter
    graph_name = f"{notebook_name}-Logseq"
    converter = LogseqConverter(logseq_output_dir, graph_name)
    
    # Process each section
    total_success = 0
    total_files = 0
    
    for section_name, section_files in sections.items():
        success, total = process_section(section_name, section_files, converter, logger, notebook_name)
        total_success += success
        total_files += total
    
    # Create main dashboard
    if total_success > 0:
        converter.create_main_dashboard()
    
    # Save image dictionary for PowerShell extraction
    if total_success > 0:
        # Always try to save image dictionary, even if empty
        dict_file = converter.save_image_dictionary(logseq_output_dir / "image_extraction_map.json")
        logger.info(f"Image dictionary saved: {dict_file}")

        if converter.image_dictionary:
            logger.info(f"Found {len(converter.image_dictionary)} images for extraction")
        else:
            logger.info("No images found in processed pages")
    
    # Create Logseq configuration file
    create_logseq_config(converter.graph_root, logger)
    
    # Summary
    log_conversion_summary(logger, total_success, total_files, sections)

    if total_success > 0:
        print(f"\nSuccessfully processed {total_success} page(s)")
        print(f"Logseq graph created at: {converter.graph_root}")
        
        # Show graph structure
        print("\nLogseq Graph Structure:")
        show_graph_structure(converter.graph_root)
        
        # Show extraction statistics
        print(f"\n📊 Extraction Statistics:")
        print(f"  - Images found: {len(converter.image_dictionary)}")
        print(f"  - Block references: {len(converter.block_references)}")
        print(f"  - Tasks detected: {len(converter.detected_tasks)}")
        print(f"  - Meetings detected: {len(converter.detected_meetings)}")
        
        # Show next steps
        if converter.image_dictionary:
            print(f"\n📷 Found {len(converter.image_dictionary)} images to extract")
            print("Next step: Run extract_images.bat to download actual images")
            print("Images will be saved to: assets/ folder")
        
        # Instructions
        print(f"\n🔮 To open in Logseq:")
        print(f"1. Open Logseq")
        print(f"2. Click 'Add graph'")
        print(f"3. Select: {converter.graph_root}")
        print(f"4. Enable 'All pages public when publishing' if needed")
        print(f"5. Start exploring with queries and the graph view!")
        
        # Show sample queries
        print(f"\n🔍 Sample Logseq Queries to Try:")
        print("{{query (task TODO)}}")
        print("{{query (page-property type [[research]])}}")
        print("{{query (between [[7 days ago]] [[today]])}}")
        print("{{query (and (tag [[#onenote-import]]) (task TODO))}}")
        
    else:
        print("\n✗ No files processed successfully")
    
    return 0 if total_success == total_files else 1

def show_graph_structure(graph_root: Path, max_depth: int = 2):
    """Display the graph structure."""
    def show_tree(path: Path, prefix: str = "", depth: int = 0):
        if depth > max_depth:
            return
            
        items = sorted([
            item for item in path.iterdir() 
            if not item.name.startswith('.') or item.name == '.logseq'
        ], key=lambda x: (x.is_file(), x.name.lower()))
        
        for i, item in enumerate(items):
            is_last = i == len(items) - 1
            current_prefix = "└── " if is_last else "├── "
            
            if item.is_dir():
                print(f"{prefix}{current_prefix}{item.name}/")
                next_prefix = prefix + ("    " if is_last else "│   ")
                # Don't recurse into .logseq directory
                if item.name != '.logseq':
                    show_tree(item, next_prefix, depth + 1)
            else:
                # Show count of files in directories
                if depth < max_depth and path.name in ['pages', 'journals']:
                    file_count = len(list(path.glob('*.md')))
                    if i == 0:  # Show count once
                        print(f"{prefix}    ({file_count} files)")
                else:
                    print(f"{prefix}{current_prefix}{item.name}")
    
    print(f"{graph_root.name}/")
    show_tree(graph_root)

def create_logseq_config(graph_root: Path, logger):
    """Create Logseq configuration."""
    config_dir = graph_root / '.logseq'
    config_dir.mkdir(exist_ok=True)
    
    # Create config.edn with basic settings
    config_content = """{
 ;; Logseq configuration for OneNote import
 
 ;; Preferred date format
 :date-formatter "MMM do, yyyy"
 
 ;; Preferred workflow
 :preferred-workflow :todo
 
 ;; Default home page
 :default-home {:page "OneNote Import Dashboard"}
 
 ;; Enable block references
 :feature/enable-block-ref? true
 
 ;; Enable rich text paste
 :rich-property-values? true
 
 ;; File name format
 :file/name-format :triple-lowbar
 
 ;; Hide specific properties in view
 :hidden-block-properties #{:id}
}"""
    
    config_path = config_dir / 'config.edn'
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(config_content)
    
    logger.info(f"Created Logseq config: {config_path}")
    
    # Create metadata.edn
    metadata_content = """{:version 1}"""
    
    metadata_path = config_dir / 'metadata.edn'
    with open(metadata_path, 'w', encoding='utf-8') as f:
        f.write(metadata_content)
    
    # Create pages-metadata.edn (empty initially)
    pages_metadata_path = config_dir / 'pages-metadata.edn'
    with open(pages_metadata_path, 'w', encoding='utf-8') as f:
        f.write("{}")

if __name__ == "__main__":
    sys.exit(main())