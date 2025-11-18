#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
Convert parsed OneNote content to markdown files
"""

from pathlib import Path
import json
import re
import html
from typing import Dict, List
import base64

class MarkdownConverter:
    """Convert parsed OneNote content to markdown."""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.image_dictionary = {}  # CallbackID -> target file path mapping
        
    def convert_section(self, parsed_data: Dict) -> Path:
        """Convert a parsed section to markdown files."""
        section_name = parsed_data['section_name']
        section_dir = self.output_dir / self._sanitize_filename(section_name)
        section_dir.mkdir(exist_ok=True)
        
        # Create assets directory
        assets_dir = section_dir / 'assets'
        assets_dir.mkdir(exist_ok=True)
        
        # Convert each page
        page_files = []
        for i, page in enumerate(parsed_data['pages'], 1):
            page_file = self._convert_page(page, i, section_dir)
            page_files.append(page_file)
        
        # Save metadata
        metadata_file = section_dir / 'section-metadata.json'
        metadata = {
            'source_file': parsed_data['source_file'],
            'section_name': section_name,
            'page_count': len(parsed_data['pages']),
            'pages': [str(f.name) for f in page_files],
            'metadata': parsed_data['metadata']
        }
        
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)
        
        print(f"Converted section '{section_name}' with {len(page_files)} pages")
        return section_dir
    
    def convert_xml_page(self, parsed_data: Dict, section_name: str) -> Path:
        """Convert a single XML page to markdown."""
        # Create section directory
        section_dir = self.output_dir / self._sanitize_filename(section_name)
        section_dir.mkdir(exist_ok=True)
        
        # Create assets directory
        assets_dir = section_dir / 'assets'
        assets_dir.mkdir(exist_ok=True)
        
        # Generate filename from page name
        page_name = parsed_data['page_name']
        safe_page_name = self._sanitize_filename(page_name)
        page_file = section_dir / f"{safe_page_name}.md"
        
        # Process images and collect for dictionary
        page_id = parsed_data.get('metadata', {}).get('ID', '')
        image_mapping = self._process_page_images(parsed_data, section_name, safe_page_name, assets_dir)
        
        # Build markdown content
        md_lines = []
        
        # Add page title
        title = parsed_data.get('title') or page_name
        md_lines.append(f"# {title}\n")
        
        # Add metadata if available
        if parsed_data.get('metadata'):
            meta = parsed_data['metadata']
            if meta.get('lastModifiedTime'):
                md_lines.append(f"*Last modified: {meta['lastModifiedTime']}*\n")
        
        # Convert content hierarchy (now with proper image links)
        for item in parsed_data['content']:
            md_content = self._convert_xml_content_item(item, image_mapping)
            if md_content:
                md_lines.append(md_content)
        
        # Write markdown file
        with open(page_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md_lines))
        
        return page_file
    
    def _process_page_images(self, parsed_data: Dict, section_name: str, 
                           page_name: str, assets_dir: Path) -> Dict[str, str]:
        """
        Process images for a page and build mapping for dictionary.
        
        Returns:
            Dict mapping CallbackID to relative image file path
        """
        page_id = parsed_data.get('metadata', {}).get('ID', '')
        image_mapping = {}
        
        # Process all images found in the page
        all_images = []
        
        # Add images from the images list
        all_images.extend(parsed_data.get('images', []))
        
        # Add images from content items
        for item in parsed_data.get('content', []):
            if item.get('type') == 'image' and isinstance(item.get('content'), dict):
                all_images.append(item['content'])
        
        # Generate file paths and add to dictionary
        for i, img in enumerate(all_images):
            callback_id = img.get('callback_id')
            if callback_id:
                # Generate image filename
                alt_text = img.get('alt', f'image_{i+1}')
                safe_alt = self._sanitize_filename(alt_text)[:20]
                
                # Create filename: section_page_description_index.ext
                image_filename = f"{section_name}_{page_name}_{safe_alt}_{i+1}.png"  # Default to PNG
                image_path = assets_dir / image_filename
                
                # Store relative path from markdown file
                relative_path = f"assets/{image_filename}"
                image_mapping[callback_id] = relative_path
                
                # Add to global dictionary for PowerShell script
                self.image_dictionary[callback_id] = {
                    'page_id': page_id,
                    'target_path': str(image_path.resolve()),
                    'relative_path': relative_path,
                    'alt_text': alt_text,
                    'section': section_name,
                    'page': page_name
                }
        
        return image_mapping
    
    def save_image_dictionary(self) -> Path:
        """Save the image dictionary to JSON file for PowerShell script."""
        dict_file = self.output_dir / 'image_extraction_map.json'
        
        with open(dict_file, 'w', encoding='utf-8') as f:
            json.dump(self.image_dictionary, f, indent=2, ensure_ascii=False)
        
        return dict_file
    
    def _convert_xml_content_item(self, item: Dict, image_mapping: Dict[str, str] = None) -> str:
        """Convert XML content item to markdown."""
        if image_mapping is None:
            image_mapping = {}
            
        item_type = item.get('type', 'unknown')
        level = item.get('level', 0)
        content = item.get('content', '')
        children = item.get('children', [])
        
        md_lines = []
        
        # Handle different content types
        if item_type == 'text' and content:
            # Convert CDATA HTML to markdown
            markdown_text = self._convert_html_to_markdown(content)
            # Add indentation for hierarchy
            indent = "  " * level
            md_lines.append(f"{indent}{markdown_text}")
        
        elif item_type == 'image':
            # Handle image content with proper links
            if isinstance(content, dict):
                callback_id = content.get('callback_id', 'unknown')
                alt_text = content.get('alt', 'Image')
                indent = "  " * level
                
                # Use proper image path if available
                if callback_id in image_mapping:
                    image_path = image_mapping[callback_id]
                    md_lines.append(f"{indent}![{alt_text}]({image_path})")
                else:
                    # Fallback to CallbackID notation
                    md_lines.append(f"{indent}![{alt_text}](CallbackID: {callback_id})")
        
        # Process children recursively
        for child in children:
            child_md = self._convert_xml_content_item(child, image_mapping)
            if child_md:
                md_lines.append(child_md)
        
        return '\n'.join(md_lines)
    
    def _convert_html_to_markdown(self, html_text: str) -> str:
        """Convert HTML in CDATA to markdown."""
        if not html_text:
            return ""
        
        # Handle highlighting (with newlines in attributes)
        html_text = re.sub(
            r"<span\s+style='background:yellow;mso-highlight:yellow'>(.*?)</span>",
            r"==\1==",
            html_text,
            flags=re.DOTALL
        )
        
        # Handle hyperlinks
        html_text = re.sub(
            r'<a href="(.*?)">(.*?)</a>',
            r'[\2](\1)',
            html_text,
            flags=re.DOTALL
        )
        
        # Handle source citations
        html_text = re.sub(
            r'From &lt;(.*?)&gt;',
            r'*Source: \1*',
            html_text
        )
        
        # Clean up remaining HTML tags (basic cleanup)
        html_text = re.sub(r'<[^>]+>', '', html_text)
        
        # Decode HTML entities (use proper decoder instead of manual replacement)
        html_text = html.unescape(html_text)
        
        return html_text.strip()
    
    def _convert_page(self, page: Dict, page_num: int, section_dir: Path) -> Path:
        """Convert a single page to markdown."""
        # Generate filename
        page_title = self._extract_page_title(page)
        filename = f"page-{page_num:03d}-{self._sanitize_filename(page_title)}.md"
        page_file = section_dir / filename
        
        # Build markdown content
        md_lines = []
        
        # Add title
        md_lines.append(f"# {page_title}\n")
        
        # Add metadata if available
        if 'date' in page:
            md_lines.append(f"*Date: {page['date']}*\n")
        
        # Convert content
        for item in page['content']:
            md_lines.append(self._convert_content_item(item))
        
        # Add image references
        if page.get('images'):
            md_lines.append("\n## Images\n")
            for i, img_src in enumerate(page['images'], 1):
                # TODO: Extract and save actual images
                md_lines.append(f"- Image {i}: `{img_src}`")
        
        # Write markdown file
        with open(page_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md_lines))
        
        return page_file
    
    def _convert_content_item(self, item: Dict) -> str:
        """Convert a content item to markdown."""
        item_type = item.get('type', 'text')
        
        if item_type == 'table':
            return self._convert_table(item.get('data', []))
        
        text = item.get('text', '').strip()
        if not text:
            return ""
        
        # Clean up text
        text = self._clean_text(text)
        
        # Apply formatting based on type
        if item_type in ['h1', 'h2', 'h3']:
            level = int(item_type[1])
            return f"{'#' * level} {text}\n"
        elif item_type == 'p':
            return f"{text}\n"
        else:
            # Default paragraph
            return f"{text}\n"
    
    def _convert_table(self, table_data: List[List[str]]) -> str:
        """Convert table data to markdown table."""
        if not table_data:
            return ""
        
        md_lines = ["\n"]
        
        # First row as header
        if len(table_data) > 0:
            header = table_data[0]
            md_lines.append("| " + " | ".join(header) + " |")
            md_lines.append("|" + "|".join([" --- " for _ in header]) + "|")
            
            # Data rows
            for row in table_data[1:]:
                # Ensure row has same number of columns
                while len(row) < len(header):
                    row.append("")
                md_lines.append("| " + " | ".join(row[:len(header)]) + " |")
        
        md_lines.append("")
        return '\n'.join(md_lines)
    
    def _extract_page_title(self, page: Dict) -> str:
        """Extract a meaningful title from page content."""
        # Use provided title if available
        if 'title' in page and page['title'] != f"Page {page.get('number', '')}":
            return page['title']
        
        # Try to find a title from content
        for item in page.get('content', [])[:5]:  # Check first 5 items
            text = item.get('text', '').strip()
            if text and len(text) < 100 and not text.startswith('http'):
                # Look for date patterns to use as title
                date_match = re.search(r'(\w+day,\s+\d{1,2}\s+\w+\s+\d{4})', text)
                if date_match:
                    return date_match.group(1)
                # Otherwise use first substantial text
                if len(text) > 10:
                    return text[:50].replace('\n', ' ')
        
        return f"Page {page.get('number', page_num)}"
    
    def _clean_text(self, text: str) -> str:
        """Clean up text content."""
        # Remove excessive whitespace
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r' {2,}', ' ', text)
        
        # Fix common encoding issues
        text = text.replace('\xa0', ' ')  # Non-breaking space
        
        # Remove redundant line breaks within paragraphs
        lines = text.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            if line:
                cleaned_lines.append(line)
        
        # Rejoin with proper spacing
        text = '\n'.join(cleaned_lines)
        
        return text
    
    def _sanitize_filename(self, name: str) -> str:
        """Sanitize string for use as filename."""
        # Remove invalid characters
        name = re.sub(r'[<>:"/\\|?*]', '-', name)
        # Limit length
        name = name[:50]
        # Remove trailing dots/spaces
        name = name.strip('. ')
        return name or 'untitled'


def main():
    """Test the markdown converter."""
    converter = MarkdownConverter(Path(__file__).parent.parent.parent / 'output')
    
    # Find parsed JSON files
    exports_dir = Path(__file__).parent.parent.parent / 'exports'
    parsed_files = list(exports_dir.glob('*_parsed.json'))
    
    for parsed_file in parsed_files:
        print(f"\nConverting: {parsed_file.name}")
        
        with open(parsed_file, 'r', encoding='utf-8') as f:
            parsed_data = json.load(f)
        
        try:
            output_dir = converter.convert_section(parsed_data)
            print(f"Output saved to: {output_dir}")
        except Exception as e:
            print(f"Error converting {parsed_file}: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()