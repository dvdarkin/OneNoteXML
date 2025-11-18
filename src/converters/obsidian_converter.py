#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
Convert parsed OneNote content to Obsidian-compatible markdown vault
"""

from pathlib import Path
import json
import re
import html
from typing import Dict, List, Optional
from datetime import datetime
import base64

from .markdown_utils import escape_literal_brackets_with_links, html_to_markdown

class ObsidianConverter:
    """Convert parsed OneNote content to Obsidian-compatible markdown vault."""
    
    def __init__(self, output_dir: Path, vault_name: str = "OneNote Vault"):
        self.vault_root = Path(output_dir) / vault_name
        self.vault_root.mkdir(parents=True, exist_ok=True)
        
        # Obsidian vault structure
        self.attachments_dir = self.vault_root / "attachments"
        self.templates_dir = self.vault_root / "templates"
        self.daily_dir = self.vault_root / "daily"
        self.references_dir = self.vault_root / "references"
        
        # Create core directories
        for dir_path in [self.attachments_dir, self.templates_dir, self.daily_dir, self.references_dir]:
            dir_path.mkdir(exist_ok=True)
            
        self.image_dictionary = {}  # CallbackID -> target file path mapping
        self.note_links = {}  # Track all notes for internal linking
        
    def convert_section(self, parsed_data: Dict) -> Path:
        """Convert a parsed section to Obsidian vault structure."""
        section_name = parsed_data['section_name']
        
        # Determine section category and placement
        section_category = self._categorize_section(section_name)
        
        # Create section index note if multiple pages
        pages = parsed_data.get('pages', [])
        if len(pages) > 1:
            self._create_section_index(section_name, pages, section_category)
        
        # Convert each page
        for page in pages:
            self._convert_page(page, section_name, section_category)
            
        return self.vault_root
        
    def _categorize_section(self, section_name: str) -> str:
        """Categorize section based on name patterns."""
        section_lower = section_name.lower()
        
        # Date-based sections (diary, journal, daily)
        if any(keyword in section_lower for keyword in ['diary', 'journal', 'daily', '2024', '2025', '2022']):
            return 'daily'
        
        # Reference materials
        if any(keyword in section_lower for keyword in ['research', 'reference', 'archive', 'papers']):
            return 'references'
        
        # Project sections
        if any(keyword in section_lower for keyword in ['project', 'tutorial', 'development']):
            return 'projects'
            
        # Default to notes
        return 'notes'
    
    def _create_section_index(self, section_name: str, pages: List[Dict], category: str):
        """Create an index note for sections with multiple pages."""
        # Create section folder
        section_folder = self._create_section_folder(section_name, category)
        
        # Create index file inside the section folder
        index_path = section_folder / "README.md"
        
        # YAML frontmatter
        frontmatter = self._generate_frontmatter(
            title=f"{section_name} Index",
            note_type="index",
            section=section_name,
            category=category,
            page_count=len(pages)
        )
        
        # Content
        content = [
            frontmatter,
            f"# {section_name}",
            "",
            f"This section contains {len(pages)} pages from OneNote.",
            "",
            "## Pages",
            ""
        ]
        
        # Add links to all pages
        for page in pages:
            page_title = page.get('title', 'Untitled')
            page_link = self._create_internal_link(page_title, section_name)
            content.append(f"- {page_link}")
            
        content.extend([
            "",
            "---",
            f"*Generated from OneNote section: {section_name}*"
        ])
        
        with open(index_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content))
            
    def _convert_page(self, page_data: Dict, section_name: str, category: str):
        """Convert a single OneNote page to Obsidian markdown."""
        # Try title first, then fall back to page_name, then default
        page_title = page_data.get('title') or page_data.get('page_name') or 'Untitled'

        # Ensure we have a string
        if not isinstance(page_title, str):
            page_title = str(page_title) if page_title is not None else 'Untitled'

        # Strip HTML tags from title (OneNote titles can contain HTML formatting)
        page_title = self._strip_html_tags(page_title)

        # Ensure title is not empty
        if not page_title.strip():
            page_title = 'Untitled'
        
        # Create filename and determine location (preserve hierarchy)
        if category == 'daily' and self._is_date_like(page_title):
            # Handle daily notes with proper date format
            filename = self._normalize_date_title(page_title)
            note_path = self.daily_dir / f"{filename}.md"
        else:
            # Create section folder and place page inside it
            section_folder = self._create_section_folder(section_name, category)
            filename = self._sanitize_filename(page_title)
            note_path = section_folder / f"{filename}.md"
        
        # Track note for internal linking (include folder for non-daily notes)
        if category == 'daily':
            self.note_links[page_title] = filename
        else:
            section_folder_name = self._sanitize_filename(section_name)
            self.note_links[page_title] = f"{section_folder_name}/{filename}"
        
        # Generate YAML frontmatter
        frontmatter = self._generate_frontmatter(
            title=page_title,
            section=section_name,
            category=category,
            last_modified=page_data.get('last_modified'),
            onenote_page_id=page_data.get('page_id')
        )
        
        # Convert content
        content_lines = [frontmatter]
        
        # Add title if not in daily notes
        if category != 'daily':
            content_lines.extend([f"# {page_title}", ""])
        
        # Convert page content
        content = page_data.get('content', [])
        for item in content:
            converted = self._convert_content_item(item, section_name, page_title)
            if converted:
                content_lines.extend(converted)
                content_lines.append("")  # Add spacing between elements
        
        # Add source attribution
        content_lines.extend([
            "---",
            f"*Extracted from OneNote: {section_name} > {page_title}*",
            f"*Last modified: {page_data.get('last_modified', 'Unknown')}*"
        ])
        
        # Write file
        with open(note_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content_lines))
            
        print(f"  Created: {note_path.name}")
        
    def _convert_content_item(self, item: Dict, section_name: str, page_title: str) -> List[str]:
        """Convert a single content item to Obsidian markdown."""
        item_type = item.get('type', 'unknown')
        
        if item_type == 'text':
            return self._convert_text(item)
        elif item_type == 'outline_element':
            return self._convert_outline_element(item, section_name, page_title)
        elif item_type == 'list':
            return self._convert_list(item)
        elif item_type == 'image':
            # For top-level image items, the actual image data is in item['content']
            image_data = item.get('content', {})
            if isinstance(image_data, dict) and image_data.get('type') == 'image':
                return self._convert_image(image_data, section_name, page_title)
            else:
                return self._convert_image(item, section_name, page_title)
        elif item_type == 'table':
            return self._convert_table(item, section_name, page_title)
        elif item_type == 'unknown_html':
            return self._convert_html_content(item)
        else:
            # Log unknown types but still try to extract any content
            if item.get('content'):
                return [f"<!-- Unknown type: {item_type} -->\n{item['content']}"]
            return [f"<!-- Unknown content type: {item_type} -->"]
    
    def _convert_text(self, item: Dict) -> List[str]:
        """Convert text content with Obsidian-specific formatting."""
        # Try 'content' first, then 'text' for backward compatibility
        text = item.get('content') or item.get('text', '')
        level = item.get('level', 0)
        
        if not text or not str(text).strip():
            return []
        
        # Ensure text is a string
        text = str(text)
        
        # Convert HTML to Obsidian markdown
        text = self._html_to_obsidian_markdown(text)
        
        # Handle headers
        if level > 0:
            header_level = min(level + 1, 6)  # Obsidian supports h1-h6
            return [f"{'#' * header_level} {text}"]
        
        return [text]
    
    def _convert_image(self, item: Dict, section_name: str, page_title: str) -> List[str]:
        """Convert image to Obsidian format with proper attachment handling."""
        callback_id = item.get('callback_id')
        # Handle both 'alt_text' and 'alt' field names from XML parser
        alt_text = item.get('alt_text') or item.get('alt', 'Image')
        
        if not callback_id:
            return [f"<!-- Missing image: {alt_text} -->"]
        
        # Generate Obsidian-compatible image filename
        image_name = self._generate_image_name(section_name, page_title, alt_text, len(self.image_dictionary))
        
        # Use Obsidian wikilink format for images
        # Option 1: Wikilink format (Obsidian native)
        obsidian_link = f"![[{image_name}]]"
        
        # Option 2: Standard markdown (for compatibility)
        # markdown_link = f"![{alt_text}](attachments/{image_name})"
        
        # Store in image dictionary for PowerShell extraction
        self.image_dictionary[callback_id] = {
            'page_id': item.get('page_id', callback_id),  # Use callback_id as fallback
            'target_path': str(self.attachments_dir / image_name),
            'relative_path': f"attachments/{image_name}",
            'alt_text': alt_text,
            'section': section_name,
            'page': page_title,
            'obsidian_link': obsidian_link
        }
        
        # Return Obsidian-formatted image
        if alt_text and alt_text.lower() != 'image':
            return [
                obsidian_link,
                f"*{alt_text}*"  # Caption below image
            ]
        else:
            return [obsidian_link]
    
    def _convert_table(self, item: Dict, section_name: str, page_title: str) -> List[str]:
        """Convert table to Obsidian markdown table with image support."""
        # Check if this is the new table format with embedded images
        if isinstance(item, dict) and item.get('type') == 'table':
            return self._convert_enhanced_table(item, section_name, page_title)
        
        # Handle legacy simple table format
        rows = item.get('rows', [])
        if not rows:
            return ["<!-- Empty table -->"]

        lines = []

        # Convert all rows through HTML-to-markdown first
        converted_rows = []
        for row in rows:
            converted_row = [self._html_to_obsidian_markdown(str(cell)) for cell in row]
            # Remove trailing empty cells
            while converted_row and not converted_row[-1].strip():
                converted_row.pop()
            converted_rows.append(converted_row)

        # Find maximum column count (excluding empty trailing cells)
        max_cols = max(len(row) for row in converted_rows) if converted_rows else 0

        if max_cols == 0:
            return ["<!-- Empty table -->"]

        # Pad all rows to have the same number of columns
        for i, row in enumerate(converted_rows):
            while len(row) < max_cols:
                row.append("")

        # Create table
        if len(converted_rows) > 0:
            # First row is header
            lines.append("| " + " | ".join(converted_rows[0]) + " |")
            lines.append("| " + " | ".join(["---"] * max_cols) + " |")

            # Remaining rows are data
            for row in converted_rows[1:]:
                lines.append("| " + " | ".join(row) + " |")

        return lines
    
    def _convert_enhanced_table(self, table_data: Dict, section_name: str, page_title: str) -> List[str]:
        """Convert enhanced table structure with embedded images."""
        lines = []
        
        # Extract the actual table content from the nested structure
        content = table_data.get('content', {})
        if not isinstance(content, dict):
            return ["<!-- Invalid table structure -->"]
        
        rows = content.get('rows', [])
        
        if not rows:
            return ["<!-- Empty table -->"]
        
        # Process embedded images first
        if content.get('has_images') and content.get('embedded_images'):
            for img in content['embedded_images']:
                # Add image to our dictionary for extraction
                img_lines = self._convert_image(img, section_name, page_title)
                # Don't add image lines directly to table, we'll reference them
        
        # Build table with proper cell content
        all_row_cells = []
        for row_idx, row in enumerate(rows):
            row_cells = []

            for cell in row:
                if isinstance(cell, dict):
                    cell_content = cell.get('text', '').strip()

                    # If cell has images, add them after the text
                    if cell.get('has_images') and cell.get('images'):
                        for img in cell['images']:
                            # Convert image and get the Obsidian link
                            img_lines = self._convert_image(img, section_name, page_title)
                            if img_lines and img_lines[0].startswith('![['):
                                # Add image link to cell content
                                if cell_content:
                                    cell_content += '<br>' + img_lines[0]
                                else:
                                    cell_content = img_lines[0]

                    # Convert cell content through HTML-to-markdown
                    row_cells.append(self._html_to_obsidian_markdown(cell_content))
                else:
                    # Legacy: plain text cell - convert through HTML-to-markdown
                    row_cells.append(self._html_to_obsidian_markdown(str(cell)))

            # Remove trailing empty cells
            while row_cells and not row_cells[-1].strip():
                row_cells.pop()

            all_row_cells.append(row_cells)

        # Find maximum column count
        max_cols = max(len(row) for row in all_row_cells) if all_row_cells else 0

        if max_cols == 0:
            return ["<!-- Empty table -->"]

        # Pad all rows to have the same number of columns
        for row_cells in all_row_cells:
            while len(row_cells) < max_cols:
                row_cells.append("")

        # Create table rows
        for row_idx, row_cells in enumerate(all_row_cells):
            lines.append("| " + " | ".join(row_cells) + " |")

            # Add separator after first row (header)
            if row_idx == 0:
                lines.append("| " + " | ".join(["---"] * max_cols) + " |")
        
        return lines
    
    def _html_to_obsidian_markdown(self, html_text: str) -> str:
        """Convert HTML to Obsidian-compatible markdown."""
        # Use shared HTML-to-markdown converter with Obsidian highlight syntax
        text = html_to_markdown(html_text, highlight_syntax='==')

        # Obsidian-specific post-processing

        # Fix: Escape literal brackets that contain markdown links
        # Pattern: [link1, link2] becomes \[link1, link2\] to avoid syntax conflict
        text = escape_literal_brackets_with_links(text)

        return text
    
    def _is_empty_html_tag(self, text: str) -> bool:
        """Check if text is just empty HTML tags like <p>, </p>, <blockquote>, etc."""
        if not text:
            return True
        
        # Remove HTML tags and check if there's actual content
        import re
        content_without_tags = re.sub(r'<[^>]+>', '', text).strip()
        return len(content_without_tags) == 0
    
    def _convert_outline_element(self, item: Dict, section_name: str, page_title: str) -> List[str]:
        """Convert outline element to Obsidian markdown."""
        lines = []
        
        # Get content from the outline element
        content = item.get('content')
        children = item.get('children', [])
        level = item.get('level', 0)
        
        # Process main content
        if content:
            if isinstance(content, str):
                # Text content - skip if it's just HTML tags with no real content
                stripped_content = content.strip()
                if stripped_content and not self._is_empty_html_tag(stripped_content):
                    converted_text = self._html_to_obsidian_markdown(content)
                    if converted_text.strip():  # Only add if conversion produces content
                        if level > 0:
                            # Use as header
                            header_level = min(level + 1, 6)
                            lines.append(f"{'#' * header_level} {converted_text}")
                        else:
                            lines.append(converted_text)
            elif isinstance(content, dict):
                # Nested content (like images)
                # Pass the full content dict to handle images properly
                if content.get('type') == 'image':
                    # For images, we need to pass the image data directly
                    converted = self._convert_image(content, section_name, page_title)
                    lines.extend(converted)
                else:
                    converted = self._convert_content_item(content, section_name, page_title)
                    lines.extend(converted)
        
        # Process children recursively
        if children:
            for child in children:
                child_lines = self._convert_content_item(child, section_name, page_title)
                if child_lines:
                    # Add indentation for nested content
                    if level > 0:
                        indented_lines = [f"  {line}" if line.strip() else line for line in child_lines]
                        lines.extend(indented_lines)
                    else:
                        lines.extend(child_lines)
        
        return lines
    
    def _generate_frontmatter(self, **properties) -> str:
        """Generate YAML frontmatter for Obsidian."""
        frontmatter_data = {
            'title': properties.get('title', 'Untitled'),
            'date': datetime.now().strftime('%Y-%m-%d'),
            'tags': ['onenote-import'],
            'onenote_source': f"{properties.get('section', '')} > {properties.get('title', '')}",
            'extraction_date': datetime.now().strftime('%Y-%m-%d')
        }
        
        # Add optional properties
        if properties.get('category'):
            frontmatter_data['category'] = properties['category']
        if properties.get('last_modified'):
            frontmatter_data['last_modified'] = properties['last_modified']
        if properties.get('onenote_page_id'):
            frontmatter_data['onenote_page_id'] = properties['onenote_page_id']
        if properties.get('note_type'):
            frontmatter_data['type'] = properties['note_type']
        if properties.get('page_count'):
            frontmatter_data['page_count'] = properties['page_count']
        
        # Format as YAML
        lines = ['---']
        for key, value in frontmatter_data.items():
            if isinstance(value, list):
                if len(value) == 1:
                    lines.append(f'{key}: [{value[0]}]')
                else:
                    lines.append(f'{key}: {value}')
            else:
                lines.append(f'{key}: {value}')
        lines.append('---')
        lines.append('')
        
        return '\n'.join(lines)
    
    def _create_section_folder(self, section_name: str, category: str) -> Path:
        """Create and return section folder path, preserving OneNote hierarchy."""
        # Sanitize section name for folder
        folder_name = self._sanitize_filename(section_name)
        
        # Determine base location based on category
        if category == 'daily':
            base_dir = self.daily_dir
        elif category == 'references':
            base_dir = self.references_dir
        else:
            # All other content goes in vault root to maintain flat-ish structure
            # but still organized by section
            base_dir = self.vault_root
        
        # Create section folder
        section_folder = base_dir / folder_name
        section_folder.mkdir(exist_ok=True)
        
        # Note: Images will still go to central attachments folder for Obsidian best practices
        # but we create the section folder structure for organization
        
        return section_folder
    
    def _strip_html_tags(self, text: str) -> str:
        """Strip HTML tags from text, keeping only the text content."""
        if not text:
            return ''

        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', text)

        # Decode HTML entities
        text = html.unescape(text)

        # Clean up excessive whitespace
        text = re.sub(r'\s+', ' ', text).strip()

        return text

    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for Obsidian (no special characters)."""
        # Handle non-string inputs
        if not isinstance(filename, str):
            filename = str(filename) if filename is not None else 'untitled'

        # Handle empty or whitespace-only strings
        if not filename.strip():
            filename = 'untitled'

        # Replace spaces with hyphens
        filename = re.sub(r'\s+', '-', filename)

        # Remove special characters, keep alphanumeric, hyphens, underscores
        filename = re.sub(r'[^\w\-_]', '', filename)

        # Remove multiple consecutive hyphens
        filename = re.sub(r'-+', '-', filename)

        # Trim hyphens from start/end
        filename = filename.strip('-_')

        # Ensure not empty
        if not filename:
            filename = 'untitled'

        return filename
    
    def _generate_image_name(self, section: str, page: str, alt_text: str, count: int) -> str:
        """Generate Obsidian-compatible image filename.

        Note: Default extension is .png. PowerShell may change extension based on
        actual image format during extraction. Extension mismatches are acceptable
        since most images are PNG anyway, and Obsidian will find them.
        """
        # Use alt_text if meaningful, otherwise use section-page-number
        if alt_text and alt_text.lower() not in ['image', 'untitled', '']:
            base_name = self._sanitize_filename(alt_text)
        else:
            base_name = self._sanitize_filename(f"{section}-{page}-{count}")

        # Default to .png (most common format)
        # PowerShell script will save with correct extension
        return f"{base_name}.png"
    
    def _create_internal_link(self, page_title: str, section_name: str) -> str:
        """Create Obsidian internal link."""
        if page_title in self.note_links:
            return f"[[{self.note_links[page_title]}|{page_title}]]"
        else:
            # Fallback to display name
            return f"[[{page_title}]]"
    
    def _is_date_like(self, title: str) -> bool:
        """Check if title looks like a date."""
        date_patterns = [
            r'\d{4}-\d{2}-\d{2}',  # YYYY-MM-DD
            r'\d{1,2}/\d{1,2}/\d{4}',  # MM/DD/YYYY
            r'(January|February|March|April|May|June|July|August|September|October|November|December)',
            r'(diary|journal)\s*\d{4}'
        ]
        
        for pattern in date_patterns:
            if re.search(pattern, title, re.IGNORECASE):
                return True
        return False
    
    def _normalize_date_title(self, title: str) -> str:
        """Convert date-like titles to YYYY-MM-DD format."""
        # This is a simple implementation - could be enhanced
        # For now, just sanitize the title
        return self._sanitize_filename(title)
    
    def _convert_list(self, item: Dict) -> List[str]:
        """Convert list to Obsidian markdown."""
        items = item.get('items', [])
        list_type = item.get('list_type', 'unordered')
        
        lines = []
        for i, list_item in enumerate(items):
            if list_type == 'ordered':
                prefix = f"{i + 1}. "
            else:
                prefix = "- "
            
            # Convert list item content
            item_text = list_item.get('text', '')
            item_text = self._html_to_obsidian_markdown(item_text)
            lines.append(f"{prefix}{item_text}")
        
        return lines
    
    def _convert_html_content(self, item: Dict) -> List[str]:
        """Convert unknown HTML content."""
        html = item.get('html', '')
        
        # Try to convert to markdown
        converted = self._html_to_obsidian_markdown(html)
        
        if converted.strip():
            return [converted]
        else:
            return ["<!-- Unknown HTML content -->"]
    
    def save_image_dictionary(self, output_path: Path = None):
        """Save image dictionary for PowerShell extraction."""
        if output_path is None:
            output_path = self.vault_root.parent / "image_extraction_map.json"
        
        # Convert paths to strings for JSON serialization
        json_dict = {}
        for callback_id, image_info in self.image_dictionary.items():
            json_dict[callback_id] = {
                'page_id': image_info['page_id'],
                'target_path': str(image_info['target_path']),
                'relative_path': image_info['relative_path'],
                'alt_text': image_info['alt_text'],
                'section': image_info['section'],
                'page': image_info['page']
            }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(json_dict, f, indent=2, ensure_ascii=False)
        
        print(f"Image extraction map saved: {output_path}")
        print(f"Found {len(self.image_dictionary)} images for extraction")
        
        return output_path