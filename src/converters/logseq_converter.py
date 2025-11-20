#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
Convert parsed OneNote content to Logseq-compatible markdown format
"""

from pathlib import Path
import json
import re
import html
from typing import Dict, List, Optional, Set
from datetime import datetime
import hashlib

from .markdown_utils import escape_literal_brackets_with_links, escape_logseq_special_syntax, html_to_markdown

class LogseqConverter:
    """Convert parsed OneNote content to Logseq-compatible markdown format."""
    
    def __init__(self, output_dir: Path, graph_name: str = "OneNote-Logseq"):
        self.graph_root = Path(output_dir) / graph_name
        self.graph_root.mkdir(parents=True, exist_ok=True)
        
        # Logseq graph structure
        self.pages_dir = self.graph_root / "pages"
        self.journals_dir = self.graph_root / "journals"
        self.assets_dir = self.graph_root / "assets"
        self.logseq_dir = self.graph_root / ".logseq"
        
        # Create core directories
        for dir_path in [self.pages_dir, self.journals_dir, self.assets_dir, self.logseq_dir]:
            dir_path.mkdir(exist_ok=True)
            
        self.image_dictionary = {}  # CallbackID -> target file path mapping
        self.block_references = {}  # ObjectID -> block reference mapping
        self.page_links = {}  # Track all pages for internal linking
        self.detected_tasks = []  # Track TODO items for dashboard
        self.detected_meetings = []  # Track meetings for dashboard
        self.used_image_names = set()  # Track used names to avoid collisions
        
    def convert_section(self, parsed_data: Dict) -> Path:
        """Convert a parsed section to Logseq format."""
        section_name = parsed_data['section_name']
        pages = parsed_data.get('pages', [])
        
        # Determine if this is a journal section
        is_journal = self._is_journal_section(section_name)
        
        # Convert each page
        for page in pages:
            self._convert_page(page, section_name, is_journal)
            
        # Create section dashboard if multiple pages
        if len(pages) > 1 and not is_journal:
            self._create_section_dashboard(section_name, pages)
            
        return self.graph_root
        
    def _is_journal_section(self, section_name: str) -> bool:
        """Check if section contains diary/journal entries."""
        journal_keywords = ['diary', 'journal', 'daily', 'log']
        return any(keyword in section_name.lower() for keyword in journal_keywords)
    
    def _convert_page(self, page_data: Dict, section_name: str, is_journal: bool):
        """Convert a single OneNote page to Logseq format."""
        page_title = page_data.get('title') or page_data.get('page_name') or 'Untitled'

        # Ensure we have a string
        if not isinstance(page_title, str):
            page_title = str(page_title) if page_title is not None else 'Untitled'

        # Strip HTML tags from title (OneNote titles can contain HTML formatting)
        page_title = self._strip_html_tags(page_title)

        # Ensure title is not empty
        if not page_title.strip():
            page_title = 'Untitled'
        
        # Check if this is a date-based page
        date_match = self._extract_date_from_content(page_data)
        
        if is_journal and date_match:
            # Create as journal page
            filename = self._format_journal_date(date_match)
            note_path = self.journals_dir / f"{filename}.md"
            is_journal_page = True
        else:
            # Create as regular page
            filename = self._sanitize_page_name(page_title)
            note_path = self.pages_dir / f"{filename}.md"
            is_journal_page = False
        
        # Track page for internal linking
        self.page_links[page_title] = filename
        
        # Convert content to Logseq format
        content_lines = []
        
        # Add page title as first block (Logseq convention)
        if not is_journal_page:
            content_lines.append(f"- # {page_title}")
        
        # Add properties block
        properties = self._generate_properties_block(
            page_data, section_name, is_journal_page
        )
        if properties:
            content_lines.extend(properties)
        
        # Convert page content
        content = page_data.get('content', [])
        for item in content:
            converted = self._convert_content_item(
                item, section_name, page_title, 
                parent_object_id=page_data.get('page_id')
            )
            if converted:
                content_lines.extend(converted)
        
        # Add source attribution as collapsed block
        content_lines.extend([
            "- {{collapsed}}",
            f"  - Extracted from OneNote: {section_name} > {page_title}",
            f"  - Last modified: {page_data.get('last_modified', 'Unknown')}"
        ])
        
        # Write file
        with open(note_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content_lines))
            
        print(f"  Created: {note_path.name}")
        
    def _convert_content_item(self, item: Dict, section_name: str, page_title: str, 
                             indent_level: int = 0, parent_object_id: str = None) -> List[str]:
        """Convert a single content item to Logseq block format."""
        item_type = item.get('type', 'unknown')
        object_id = item.get('object_id')
        
        # Generate block reference if we have an object ID
        if object_id:
            block_ref = self._generate_block_reference(object_id)
            self.block_references[object_id] = block_ref
        else:
            block_ref = None
        
        lines = []
        indent = "  " * indent_level
        
        if item_type == 'text':
            text_lines = self._convert_text(item, indent_level)
            if text_lines and block_ref:
                # Add block reference to first line
                text_lines[0] = text_lines[0].replace("- ", f"- {block_ref} ", 1)
            lines.extend(text_lines)
            
        elif item_type == 'outline_element':
            oe_lines = self._convert_outline_element(
                item, section_name, page_title, indent_level, parent_object_id
            )
            lines.extend(oe_lines)
            
        elif item_type == 'list':
            list_lines = self._convert_list(item, indent_level)
            lines.extend(list_lines)
            
        elif item_type == 'image':
            image_data = item.get('content', {})
            if isinstance(image_data, dict) and image_data.get('type') == 'image':
                image_lines = self._convert_image(image_data, section_name, page_title, indent_level)
            else:
                image_lines = self._convert_image(item, section_name, page_title, indent_level)
            lines.extend(image_lines)
            
        elif item_type == 'table':
            table_lines = self._convert_table(item, section_name, page_title, indent_level)
            lines.extend(table_lines)
            
        elif item_type == 'unknown_html':
            html_lines = self._convert_html_content(item, indent_level)
            lines.extend(html_lines)
            
        else:
            # Log unknown types but still try to extract any content
            if item.get('content'):
                lines.append(f"{indent}- <!--Unknown type: {item_type}-->")
                lines.append(f"{indent}  {item['content']}")
            else:
                lines.append(f"{indent}- <!--Unknown content type: {item_type}-->")
        
        return lines
    
    def _convert_text(self, item: Dict, indent_level: int = 0) -> List[str]:
        """Convert text content to Logseq block format."""
        text = item.get('content') or item.get('text', '')
        level = item.get('level', 0)
        
        if not text or not str(text).strip():
            return []
        
        # Ensure text is a string
        text = str(text)
        
        # Convert HTML to Logseq markdown
        text = self._html_to_logseq_markdown(text)
        
        # Detect special content types
        self._detect_special_content(text, item)
        
        indent = "  " * indent_level
        
        # Always format as block (Logseq convention)
        if level > 0:
            # Headers in Logseq are just blocks with # prefix
            header_level = min(level + 1, 6)
            return [f"{indent}- {'#' * header_level} {text}"]
        else:
            return [f"{indent}- {text}"]
    
    def _convert_outline_element(self, item: Dict, section_name: str, page_title: str, 
                                indent_level: int = 0, parent_object_id: str = None) -> List[str]:
        """Convert outline element to Logseq blocks."""
        lines = []
        content = item.get('content')
        children = item.get('children', [])
        object_id = item.get('object_id', parent_object_id)
        
        # Process main content
        if content:
            if isinstance(content, str):
                stripped_content = content.strip()
                if stripped_content and not self._is_empty_html_tag(stripped_content):
                    converted_text = self._html_to_logseq_markdown(content)
                    if converted_text.strip():
                        indent = "  " * indent_level
                        
                        # Add block reference if available
                        if object_id and object_id not in self.block_references:
                            block_ref = self._generate_block_reference(object_id)
                            self.block_references[object_id] = block_ref
                            lines.append(f"{indent}- {block_ref} {converted_text}")
                        else:
                            lines.append(f"{indent}- {converted_text}")
                            
                        # Detect special content
                        self._detect_special_content(converted_text, item)
                        
            elif isinstance(content, dict):
                # Nested content
                if content.get('type') == 'image':
                    converted = self._convert_image(content, section_name, page_title, indent_level)
                else:
                    converted = self._convert_content_item(
                        content, section_name, page_title, indent_level, object_id
                    )
                lines.extend(converted)
        
        # Process children as nested blocks
        if children:
            for child in children:
                child_lines = self._convert_content_item(
                    child, section_name, page_title, indent_level + 1, object_id
                )
                lines.extend(child_lines)
        
        return lines
    
    def _convert_image(self, item: Dict, section_name: str, page_title: str, 
                      indent_level: int = 0) -> List[str]:
        """Convert image to Logseq format."""
        callback_id = item.get('callback_id')
        alt_text = item.get('alt_text') or item.get('alt', 'Image')
        
        indent = "  " * indent_level
        
        if not callback_id:
            return [f"{indent}- <!--Missing image: {alt_text}-->"]
        
        # Generate Logseq-compatible image filename  
        # Try to detect format from callback_id or alt_text hints
        image_format = self._detect_image_format_hint(item)
        image_name = self._generate_image_name(section_name, page_title, alt_text, len(self.image_dictionary), image_format)
        
        # Logseq uses standard markdown for images
        markdown_link = f"![{alt_text}](../assets/{image_name})"
        
        # Store in image dictionary for PowerShell extraction
        # Generate original section name by reversing common sanitizations
        original_section_name = self._reverse_sanitize_section_name(section_name)
        
        self.image_dictionary[callback_id] = {
            'page_id': item.get('page_id', callback_id),
            'target_path': str(self.assets_dir / image_name),
            'relative_path': f"assets/{image_name}",
            'alt_text': alt_text,
            'section': original_section_name,  # Original name for PowerShell lookup
            'section_sanitized': section_name,  # Sanitized name for reference
            'page': page_title,
            'logseq_link': markdown_link
        }
        
        # Return as block with image
        lines = [f"{indent}- {markdown_link}"]
        
        # Add caption as nested block if meaningful
        if alt_text and alt_text.lower() not in ['image', 'untitled', '']:
            lines.append(f"{indent}  - *{alt_text}*")
            
        return lines
    
    def _convert_table(self, item: Dict, section_name: str, page_title: str, 
                      indent_level: int = 0) -> List[str]:
        """Convert table to Logseq format."""
        indent = "  " * indent_level
        lines = []
        
        # Add table as a block
        lines.append(f"{indent}- Table:")
        
        # Check if this is the new table format with embedded images
        if isinstance(item, dict) and item.get('type') == 'table':
            table_lines = self._convert_enhanced_table(item, section_name, page_title, indent_level + 1)
        else:
            # Handle legacy simple table format
            table_lines = self._convert_simple_table(item, indent_level + 1)
        
        lines.extend(table_lines)
        return lines
    
    def _convert_simple_table(self, item: Dict, indent_level: int) -> List[str]:
        """Convert simple table format."""
        rows = item.get('rows', [])
        if not rows:
            return ["  - <!--Empty table-->"]

        lines = []
        indent = "  " * indent_level

        # Convert all rows and remove trailing empty cells
        all_converted_rows = []
        for row in rows:
            converted_cells = [self._html_to_logseq_markdown(str(cell)) for cell in row]
            # Remove trailing empty cells
            while converted_cells and not converted_cells[-1].strip():
                converted_cells.pop()
            all_converted_rows.append(converted_cells)

        # Find maximum column count
        max_cols = max(len(row) for row in all_converted_rows) if all_converted_rows else 0

        if max_cols == 0:
            return ["  - <!--Empty table-->"]

        # Pad all rows to have the same number of columns
        for row_cells in all_converted_rows:
            while len(row_cells) < max_cols:
                row_cells.append("")

        # Convert to nested list format (Logseq doesn't have native tables)
        for row_idx, converted_cells in enumerate(all_converted_rows):
            if row_idx == 0:
                # Header row
                lines.append(f"{indent}- **{' | '.join(converted_cells)}**")
            else:
                # Data row
                lines.append(f"{indent}- {' | '.join(converted_cells)}")

        return lines
    
    def _convert_enhanced_table(self, table_data: Dict, section_name: str, 
                               page_title: str, indent_level: int) -> List[str]:
        """Convert enhanced table with embedded images."""
        lines = []
        indent = "  " * indent_level
        
        content = table_data.get('content', {})
        if not isinstance(content, dict):
            return [f"{indent}- <!--Invalid table structure-->"]
        
        rows = content.get('rows', [])
        if not rows:
            return [f"{indent}- <!--Empty table-->"]
        
        # Process embedded images first
        if content.get('has_images') and content.get('embedded_images'):
            for img in content['embedded_images']:
                self._convert_image(img, section_name, page_title, 0)
        
        # Build table as nested blocks
        all_row_cells = []
        for row_idx, row in enumerate(rows):
            row_cells = []

            for cell in row:
                if isinstance(cell, dict):
                    cell_content = cell.get('text', '').strip()

                    # If cell has images, add them
                    if cell.get('has_images') and cell.get('images'):
                        for img in cell['images']:
                            img_lines = self._convert_image(img, section_name, page_title, 0)
                            if img_lines and '![' in img_lines[0]:
                                # Extract just the image link
                                img_match = re.search(r'!\[.*?\]\(.*?\)', img_lines[0])
                                if img_match:
                                    if cell_content:
                                        cell_content += ' ' + img_match.group()
                                    else:
                                        cell_content = img_match.group()

                    # Convert cell content through HTML-to-markdown
                    row_cells.append(self._html_to_logseq_markdown(cell_content))
                else:
                    # Convert cell content through HTML-to-markdown
                    row_cells.append(self._html_to_logseq_markdown(str(cell)))

            # Remove trailing empty cells
            while row_cells and not row_cells[-1].strip():
                row_cells.pop()

            all_row_cells.append(row_cells)

        # Find maximum column count
        max_cols = max(len(row) for row in all_row_cells) if all_row_cells else 0

        if max_cols == 0:
            return [f"{indent}- <!--Empty table-->"]

        # Pad all rows to have the same number of columns
        for row_cells in all_row_cells:
            while len(row_cells) < max_cols:
                row_cells.append("")

        # Create rows as blocks
        for row_idx, row_cells in enumerate(all_row_cells):
            if row_idx == 0:
                # Header row
                lines.append(f"{indent}- **{' | '.join(row_cells)}**")
            else:
                lines.append(f"{indent}- {' | '.join(row_cells)}")
        
        return lines
    
    def _convert_list(self, item: Dict, indent_level: int = 0) -> List[str]:
        """Convert list to Logseq blocks."""
        items = item.get('items', [])
        list_type = item.get('list_type', 'unordered')
        
        lines = []
        indent = "  " * indent_level
        
        for i, list_item in enumerate(items):
            # Convert list item content
            item_text = list_item.get('text', '')
            item_text = self._html_to_logseq_markdown(item_text)
            
            # Check for TODO patterns
            is_task = self._is_task_item(item_text)
            
            if is_task:
                # Convert to Logseq task format
                task_state, task_text = self._parse_task_item(item_text)
                if task_state == 'TODO':
                    lines.append(f"{indent}- TODO {task_text}")
                    self.detected_tasks.append({'text': task_text, 'state': 'TODO'})
                elif task_state == 'DONE':
                    lines.append(f"{indent}- DONE {task_text}")
                    self.detected_tasks.append({'text': task_text, 'state': 'DONE'})
                else:
                    lines.append(f"{indent}- {item_text}")
            else:
                # Regular list item
                lines.append(f"{indent}- {item_text}")
        
        return lines
    
    def _convert_html_content(self, item: Dict, indent_level: int = 0) -> List[str]:
        """Convert unknown HTML content."""
        html = item.get('html', '')
        indent = "  " * indent_level
        
        # Try to convert to markdown
        converted = self._html_to_logseq_markdown(html)
        
        if converted.strip():
            return [f"{indent}- {converted}"]
        else:
            return [f"{indent}- <!--Unknown HTML content-->"]

    def _html_to_logseq_markdown(self, html_text: str) -> str:
        """Convert HTML to Logseq-compatible markdown."""
        # Use shared HTML-to-markdown converter with Logseq highlight syntax
        text = html_to_markdown(html_text, highlight_syntax='^^')

        # Logseq-specific post-processing

        # Extract dates from highlighted text for linking
        highlighted_dates = re.findall(r"\^\^(.*?)\^\^", text)
        for date_text in highlighted_dates:
            date_link = self._try_parse_date_to_link(date_text)
            if date_link:
                text = text.replace(f"^^{date_text}^^", f"^^{date_link}^^")

        # Fix: Escape literal brackets that contain markdown links
        # Pattern: [link1, link2] becomes \[link1, link2\] to avoid syntax conflict
        text = escape_literal_brackets_with_links(text)

        # Fix: Escape Logseq-specific syntax in non-code contexts
        # Prevents {{queries}}, ::properties, and ((block-refs)) from triggering in prose
        text = escape_logseq_special_syntax(text)

        return text
    
    def _generate_properties_block(self, page_data: Dict, section_name: str, 
                                  is_journal: bool) -> List[str]:
        """Generate Logseq properties for the page."""
        properties = []
        
        # Add notebook and section properties
        properties.append(f"  notebook:: [[{page_data.get('notebook_name', 'OneNote')}]]")
        properties.append(f"  section:: [[{section_name}]]")
        
        # Add temporal properties
        if page_data.get('last_modified'):
            modified_date = self._format_date_property(page_data['last_modified'])
            if modified_date:
                properties.append(f"  modified:: {modified_date}")
        
        # Add content type detection
        content_type = self._detect_content_type(page_data)
        if content_type:
            properties.append(f"  type:: [[{content_type}]]")
        
        # Add OneNote metadata
        if page_data.get('page_id'):
            properties.append(f"  onenote-id:: {page_data['page_id']}")
        
        # Add author if available
        if page_data.get('author'):
            properties.append(f"  author:: [[{page_data['author']}]]")
        
        # Add tags
        tags = self._generate_tags(section_name, content_type)
        if tags:
            properties.append(f"  tags:: {', '.join(tags)}")
        
        return ["- Properties:"] + properties if properties else []
    
    def _detect_content_type(self, page_data: Dict) -> Optional[str]:
        """Detect the type of content in the page."""
        content_str = json.dumps(page_data.get('content', []))
        
        # Check for different content patterns
        if re.search(r'(meeting|minutes|agenda)', content_str, re.IGNORECASE):
            return 'meeting-notes'
        elif re.search(r'(todo|task|action item)', content_str, re.IGNORECASE):
            return 'task-list'
        elif re.search(r'(research|analysis|study)', content_str, re.IGNORECASE):
            return 'research'
        elif re.search(r'(diary|journal|daily)', content_str, re.IGNORECASE):
            return 'diary-entry'
        elif re.search(r'(project|development|implementation)', content_str, re.IGNORECASE):
            return 'project-notes'
        
        return None
    
    def _detect_special_content(self, text: str, item: Dict):
        """Detect special content like tasks, meetings, etc."""
        # Detect TODO items
        if re.search(r'\b(TODO|TASK|Action Item)\b', text, re.IGNORECASE):
            priority = self._detect_priority(text)
            self.detected_tasks.append({
                'text': text,
                'priority': priority,
                'object_id': item.get('object_id')
            })
        
        # Detect meetings
        if re.search(r'\b(meeting|agenda|minutes)\b', text, re.IGNORECASE):
            self.detected_meetings.append({
                'text': text,
                'object_id': item.get('object_id')
            })
    
    def _detect_priority(self, text: str) -> str:
        """Detect priority level from text."""
        if re.search(r'\b(urgent|critical|high priority|asap)\b', text, re.IGNORECASE):
            return 'A'
        elif re.search(r'\b(important|medium priority)\b', text, re.IGNORECASE):
            return 'B'
        else:
            return 'C'
    
    def _is_task_item(self, text: str) -> bool:
        """Check if text represents a task."""
        task_patterns = [
            r'^\s*\[[ xX]\]',  # Checkbox format
            r'^\s*(TODO|DONE|TASK):?\s',  # TODO/DONE prefix
            r'^\s*[-*]\s*(TODO|DONE|TASK):?\s',  # List with TODO
        ]
        
        for pattern in task_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def _parse_task_item(self, text: str) -> tuple:
        """Parse task item to extract state and text."""
        # Check for checkbox
        if match := re.match(r'^\s*\[([ xX])\]\s*(.+)$', text):
            state = 'DONE' if match.group(1).lower() == 'x' else 'TODO'
            return state, match.group(2)
        
        # Check for TODO/DONE prefix
        if match := re.match(r'^\s*(?:[-*]\s*)?(TODO|DONE|TASK):?\s*(.+)$', text, re.IGNORECASE):
            state = 'DONE' if match.group(1).upper() == 'DONE' else 'TODO'
            return state, match.group(2)
        
        return 'TODO', text
    
    def _extract_date_from_content(self, page_data: Dict) -> Optional[str]:
        """Extract date from page content, especially from highlighted dates."""
        content_str = json.dumps(page_data.get('content', []))
        
        # Look for highlighted dates (converted to ^^date^^)
        highlighted_dates = re.findall(r'\^\^([^^\n]+)\^\^', content_str)
        
        for date_text in highlighted_dates:
            parsed_date = self._try_parse_date(date_text)
            if parsed_date:
                return parsed_date
        
        # Also check page title
        title = page_data.get('title', '')
        return self._try_parse_date(title)
    
    def _try_parse_date(self, text: str) -> Optional[str]:
        """Try to parse various date formats."""
        # Handle None or empty text
        if not text:
            return None
            
        # Ensure text is a string
        text = str(text)
        
        # Common date patterns
        patterns = [
            (r'(\d{1,2})\s+(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})', 
             lambda m: f"{m.group(3)}-{self._month_to_number(m.group(2))}-{m.group(1).zfill(2)}"),
            (r'(\d{4})-(\d{1,2})-(\d{1,2})', 
             lambda m: f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"),
            (r'(\d{1,2})/(\d{1,2})/(\d{4})', 
             lambda m: f"{m.group(3)}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}")
        ]
        
        for pattern, formatter in patterns:
            if match := re.search(pattern, text, re.IGNORECASE):
                try:
                    return formatter(match)
                except:
                    continue
        
        return None
    
    def _try_parse_date_to_link(self, text: str) -> Optional[str]:
        """Parse date and format as Logseq date link."""
        parsed = self._try_parse_date(text)
        if parsed:
            # Convert to Logseq date format: [[Aug 12th, 2024]]
            try:
                date_obj = datetime.strptime(parsed, '%Y-%m-%d')
                day = date_obj.day
                suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
                return f"[[{date_obj.strftime('%b')} {day}{suffix}, {date_obj.year}]]"
            except:
                return None
        return None
    
    def _format_date_property(self, date_str: str) -> Optional[str]:
        """Format date for Logseq property."""
        parsed = self._try_parse_date(date_str)
        if parsed:
            return self._try_parse_date_to_link(parsed)
        return None
    
    def _format_journal_date(self, date_str: str) -> str:
        """Format date for journal filename (YYYY_MM_DD)."""
        parsed = self._try_parse_date(date_str)
        if parsed:
            return parsed.replace('-', '_')
        # Fallback to sanitized date string
        return self._sanitize_page_name(date_str)
    
    def _month_to_number(self, month_name: str) -> str:
        """Convert month name to number."""
        months = {
            'january': '01', 'february': '02', 'march': '03', 'april': '04',
            'may': '05', 'june': '06', 'july': '07', 'august': '08',
            'september': '09', 'october': '10', 'november': '11', 'december': '12'
        }
        return months.get(month_name.lower(), '01')
    
    def _generate_tags(self, section_name: str, content_type: Optional[str]) -> List[str]:
        """Generate tags based on section and content type."""
        tags = ['#onenote-import']
        
        # Add section-based tags
        section_lower = section_name.lower()
        if 'research' in section_lower:
            tags.append('#research')
        if 'diary' in section_lower or 'journal' in section_lower:
            tags.append('#journal')
        if 'project' in section_lower:
            tags.append('#project')
        
        # Add content type tag
        if content_type:
            tags.append(f'#{content_type.replace("-", "_")}')
        
        return tags
    
    def _create_section_dashboard(self, section_name: str, pages: List[Dict]):
        """Create a dashboard page for the section with queries."""
        dashboard_name = f"{section_name} Dashboard"
        dashboard_path = self.pages_dir / f"{self._sanitize_page_name(dashboard_name)}.md"
        
        content = [
            f"- # {dashboard_name}",
            "- Properties:",
            f"  type:: [[dashboard]]",
            f"  section:: [[{section_name}]]",
            "",
            "- ## Overview",
            f"  - This section contains **{len(pages)}** pages from OneNote",
            f"  - Imported on [[{datetime.now().strftime('%b %d')}th, {datetime.now().year}]]",
            "",
            "- ## All Pages",
            "  - {{query (and (page-property section [[" + section_name + "]]) (not (page-property type [[dashboard]]))}}"
            "",
            "- ## Recent Updates", 
            "  - {{query (and (page-property section [[" + section_name + "]]) (between [[7 days ago]] [[today]])}}",
            "",
            "- ## Tasks",
            "  - {{query (and (page-property section [[" + section_name + "]]) (task TODO))}}",
            "",
            "- ## Completed Tasks",
            "  - {{query (and (page-property section [[" + section_name + "]]) (task DONE))}}",
        ]
        
        # Add specific queries based on content type
        if 'research' in section_name.lower():
            content.extend([
                "",
                "- ## Research Topics",
                "  - {{query (and (page-property section [[" + section_name + "]]) (page-property type [[research]]))}}",
            ])
        
        if any(word in section_name.lower() for word in ['diary', 'journal']):
            content.extend([
                "",
                "- ## This Week's Entries",
                "  - {{query (and (page-property section [[" + section_name + "]]) (between [[7 days ago]] [[today]]))}}",
                "",
                "- ## This Month's Entries", 
                "  - {{query (and (page-property section [[" + section_name + "]]) (between [[30 days ago]] [[today]]))}}",
            ])
        
        with open(dashboard_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content))
        
        print(f"  Created dashboard: {dashboard_path.name}")
    
    def _generate_block_reference(self, object_id: str) -> str:
        """Generate a unique block reference ID."""
        # Use first 8 chars of object_id hash for brevity
        block_id = hashlib.md5(object_id.encode()).hexdigest()[:8]
        return f"#^{block_id}"
    
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

    def _sanitize_page_name(self, name: str) -> str:
        """Sanitize page name for Logseq."""
        if not isinstance(name, str):
            name = str(name) if name is not None else 'untitled'

        if not name.strip():
            name = 'untitled'

        # Logseq doesn't like certain characters
        # Keep spaces but remove other special chars
        name = re.sub(r'[<>:"/\\|?*\[\]]', '', name)
        
        # Remove multiple spaces
        name = re.sub(r'\s+', ' ', name)
        
        # Trim
        name = name.strip()
        
        if not name:
            name = 'untitled'
            
        return name


    def _shorten_name(self, name: str, max_length: int = 10) -> str:
        """
        Intelligently shorten a name while preserving readability.

        Strategy:
        1. Remove common words (the, a, an, of, for, with, etc.)
        2. Use abbreviations for long words
        3. Limit to max_length characters
        """
        # Clean and normalize
        name = name.strip().lower()

        # Remove common words
        common_words = {'the', 'a', 'an', 'of', 'for', 'with', 'and', 'or', 'in', 'on', 'at', 'to'}
        words = name.split()
        words = [w for w in words if w not in common_words]

        if not words:
            words = name.split()  # Fallback to original

        # Join and clean
        shortened = '-'.join(words)
        shortened = re.sub(r'[^a-z0-9-]', '', shortened)  # Keep only alphanumeric and hyphens
        shortened = re.sub(r'-+', '-', shortened)  # Collapse multiple hyphens

        # Truncate if still too long
        if len(shortened) > max_length:
            # Try to keep meaningful parts
            parts = shortened.split('-')
            result = []
            current_len = 0

            for part in parts:
                if current_len + len(part) + 1 <= max_length:
                    result.append(part)
                    current_len += len(part) + 1
                else:
                    # Add abbreviated form if space
                    abbrev = part[:max(1, max_length - current_len - 1)]
                    if abbrev:
                        result.append(abbrev)
                    break

            shortened = '-'.join(result) if result else shortened[:max_length]

        return shortened.strip('-') or 'img'

    def _generate_image_name(self, section: str, page: str, alt_text: str, count: int, image_format: str = None) -> str:
        """
        Generate short, unique Logseq-compatible image filename.

        Format: {section_short}-{page_short}-{counter}.{ext}
        Example: ctrade-ideas-001.png
        Max length: ~30 characters (prevents Windows path limit issues)

        If collision detected, adds 4-char hash suffix: ctrade-ideas-001-a3f9.png
        """
        import hashlib

        # Shorten section and page names
        section_short = self._shorten_name(section, max_length=8)
        page_short = self._shorten_name(page, max_length=8)

        # Generate base filename with counter
        counter = f"{count:03d}"
        base_name = f"{section_short}-{page_short}-{counter}"

        # Determine proper file extension
        extension = self._get_image_extension(image_format)
        filename = f"{base_name}.{extension}"

        # Check for collision
        if filename in self.used_image_names:
            # Generate a 4-character hash from the full names for uniqueness
            hash_input = f"{section}-{page}-{count}".encode()
            short_hash = hashlib.md5(hash_input).hexdigest()[:4]
            filename = f"{base_name}-{short_hash}.{extension}"

        # Track this filename
        self.used_image_names.add(filename)

        return filename
    
    def _get_image_extension(self, image_format: str = None) -> str:
        """Get proper image file extension based on format."""
        if not image_format:
            return "png"  # Default fallback
        
        format_lower = image_format.lower()
        if format_lower in ['jpeg', 'jpg']:
            return "jpg"
        elif format_lower == 'png':
            return "png"
        elif format_lower == 'gif':
            return "gif"
        elif format_lower == 'bmp':
            return "bmp"
        elif format_lower == 'webp':
            return "webp"
        else:
            return "png"  # Default fallback
    
    def _detect_image_format_hint(self, item: Dict) -> Optional[str]:
        """Try to detect image format from available hints."""
        # Check alt text for format hints
        alt_text = item.get('alt_text') or item.get('alt', '')
        if alt_text:
            alt_lower = alt_text.lower()
            if any(fmt in alt_lower for fmt in ['.jpg', '.jpeg', 'jpeg']):
                return "jpeg"
            elif '.png' in alt_lower or 'png' in alt_lower:
                return "png"
            elif '.gif' in alt_lower or 'gif' in alt_lower:
                return "gif"
            elif '.bmp' in alt_lower or 'bmp' in alt_lower:
                return "bmp"
            elif '.webp' in alt_lower or 'webp' in alt_lower:
                return "webp"
        
        # Could also check filename patterns in callback_id if available
        callback_id = item.get('callback_id', '')
        if callback_id and isinstance(callback_id, str):
            if 'jpg' in callback_id.lower() or 'jpeg' in callback_id.lower():
                return "jpeg"
            elif 'png' in callback_id.lower():
                return "png"
        
        # No hint found - PowerShell will detect from binary data
        return None
    
    def _reverse_sanitize_section_name(self, sanitized_name: str) -> str:
        r"""Attempt to reverse common sanitizations to get original section name.
        
        The PowerShell export script uses: $section.name -replace '[^\w\s-]', '_'
        This means dots, periods, and special chars become underscores.
        Common patterns:
        - 'trade_research' -> 'trade.research' (dots to underscores)  
        - 'LensTutorial_2' -> 'LensTutorial 2' (spaces to underscores)
        """
        # Special known mappings for common cases
        known_mappings = {
            'trade_research': 'trade.research',
            'LensTutorial_2': 'LensTutorial 2'
        }
        
        if sanitized_name in known_mappings:
            return known_mappings[sanitized_name]
        
        # General heuristic: if name contains underscores but no dots or spaces,
        # try converting underscores to dots (common pattern for domain-style names)
        if '_' in sanitized_name and '.' not in sanitized_name and ' ' not in sanitized_name:
            # For patterns like 'trade_research' -> 'trade.research'
            if sanitized_name.count('_') == 1 and sanitized_name.replace('_', '').replace('.', '').isalpha():
                return sanitized_name.replace('_', '.')
        
        # If it contains digits and underscore, likely space was converted
        # Pattern: 'LensTutorial_2' -> 'LensTutorial 2'
        if '_' in sanitized_name and any(c.isdigit() for c in sanitized_name):
            return sanitized_name.replace('_', ' ')
        
        # Default: return as-is (no clear transformation pattern)
        return sanitized_name
    
    def _is_empty_html_tag(self, text: str) -> bool:
        """Check if text is just empty HTML tags."""
        if not text:
            return True
        
        content_without_tags = re.sub(r'<[^>]+>', '', text).strip()
        return len(content_without_tags) == 0
    
    def save_image_dictionary(self, output_path: Path = None):
        """Save image dictionary for PowerShell extraction."""
        if output_path is None:
            output_path = self.graph_root.parent / "logseq_image_extraction_map.json"
        
        # Convert paths to strings for JSON
        json_dict = {}
        for callback_id, image_info in self.image_dictionary.items():
            json_dict[callback_id] = {
                'page_id': image_info['page_id'],
                'target_path': str(image_info['target_path']),
                'relative_path': image_info['relative_path'],
                'alt_text': image_info['alt_text'],
                'section': image_info['section'],  # Original section name for PowerShell lookup
                'section_sanitized': image_info.get('section_sanitized', image_info['section']),  # Sanitized for reference
                'page': image_info['page']
            }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(json_dict, f, indent=2, ensure_ascii=False)
        
        print(f"Image extraction map saved: {output_path}")
        print(f"Found {len(self.image_dictionary)} images for extraction")
        
        # Also save block references if any
        if self.block_references:
            block_ref_path = self.graph_root / "block_references.json"
            with open(block_ref_path, 'w', encoding='utf-8') as f:
                json.dump(self.block_references, f, indent=2)
            print(f"Block references saved: {block_ref_path}")
        
        return output_path
    
    def create_main_dashboard(self):
        """Create main dashboard with queries for the entire graph."""
        dashboard_path = self.pages_dir / "OneNote Import Dashboard.md"
        
        content = [
            "- # OneNote Import Dashboard",
            "- Properties:",
            "  type:: [[dashboard]]",
            "  created:: [[" + datetime.now().strftime('%b %d') + "th, " + str(datetime.now().year) + "]]",
            "",
            "- ## Import Summary",
            f"  - Total images found: **{len(self.image_dictionary)}**",
            f"  - Tasks detected: **{len(self.detected_tasks)}**",
            f"  - Block references created: **{len(self.block_references)}**",
            "",
            "- ## All OneNote Content",
            "  - {{query (tag [[#onenote-import]])}}",
            "",
            "- ## Active Tasks",
            "  - {{query (and (tag [[#onenote-import]]) (task TODO))}}",
            "",
            "- ## Recent Imports",
            "  - {{query (and (tag [[#onenote-import]]) (between [[7 days ago]] [[today]])}}",
            "",
            "- ## By Content Type",
            "  - ### Research",
            "    {{query (and (tag [[#onenote-import]]) (page-property type [[research]])}}",
            "  - ### Meetings", 
            "    {{query (and (tag [[#onenote-import]]) (page-property type [[meeting-notes]])}}",
            "  - ### Projects",
            "    {{query (and (tag [[#onenote-import]]) (page-property type [[project-notes]])}}",
            "  - ### Journal Entries",
            "    {{query (and (tag [[#onenote-import]]) (page-property type [[diary-entry]])}}",
            "",
            "- ## Search by Section",
            "  - Use queries like:",
            "    - `{{query (page-property section [[Your Section Name]])}}`",
        ]
        
        with open(dashboard_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content))
        
        print(f"Created main dashboard: {dashboard_path.name}")