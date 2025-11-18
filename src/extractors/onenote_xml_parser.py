#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNote XML Parser - Enhanced for OneNote 2013 XML structure
Parse OneNote page XML exports from PowerShell COM automation
Handles OE hierarchies, CDATA content, and unknown pattern detection
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Any
import json
from datetime import datetime
import re


class OneNoteXMLParser:
    """Parse OneNote page XML content with comprehensive structure handling."""
    
    def __init__(self):
        self.namespace = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
        self.unknown_elements = set()
        self.unknown_attributes = set()
        self.style_definitions = {}
    
    def parse_page_xml(self, xml_path: Path) -> Dict[str, Any]:
        """Parse a OneNote page XML file with enhanced structure analysis."""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Validate OneNote page format
            if not root.tag.endswith('Page'):
                raise ValueError(f"Not a OneNote page XML: {root.tag}")
            
            # Reset parser state
            self.unknown_elements.clear()
            self.unknown_attributes.clear()
            self.style_definitions.clear()
            
            # Extract page metadata
            page_metadata = self._extract_page_metadata(root)
            
            # Extract style definitions first
            self.style_definitions = self._extract_style_definitions(root)
            
            # Extract page title
            title = self._extract_title(root)
            
            # Extract content with hierarchy preservation
            content = self._extract_content_hierarchy(root)
            
            # Extract images with CallbackID mapping
            images = self._extract_images(root)
            
            page_info = {
                'source_file': str(xml_path),
                'page_name': page_metadata.get('name', 'Untitled'),
                'title': title,
                'metadata': page_metadata,
                'content': content,
                'images': images,
                'style_definitions': self.style_definitions,
                'unknown_elements': list(self.unknown_elements),
                'unknown_attributes': list(self.unknown_attributes),
                'parsing_stats': {
                    'content_items': len(content),
                    'image_count': len(images),
                    'style_count': len(self.style_definitions),
                    'unknown_element_count': len(self.unknown_elements),
                    'unknown_attribute_count': len(self.unknown_attributes)
                }
            }
            
            return page_info
            
        except ET.ParseError as e:
            raise ValueError(f"Invalid XML in {xml_path}: {e}")
        except Exception as e:
            raise ValueError(f"Error parsing {xml_path}: {e}")
    
    def _extract_page_metadata(self, root) -> Dict[str, Any]:
        """Extract comprehensive metadata from Page element."""
        metadata = {}
        
        # Known page attributes
        known_attrs = {'ID', 'name', 'dateTime', 'lastModifiedTime', 'pageLevel', 'lang'}
        
        for attr, value in root.attrib.items():
            if attr in known_attrs:
                if attr in ['dateTime', 'lastModifiedTime']:
                    try:
                        metadata[attr] = self._parse_datetime(value)
                    except:
                        metadata[attr] = value
                else:
                    metadata[attr] = value
            else:
                self.unknown_attributes.add(f"Page.{attr}")
        
        # Extract page settings
        page_settings = root.find('.//one:PageSettings', self.namespace)
        if page_settings is not None:
            settings = {}
            for attr, value in page_settings.attrib.items():
                settings[attr] = value
            metadata['page_settings'] = settings
        
        return metadata
    
    def _extract_style_definitions(self, root) -> Dict[str, Dict[str, str]]:
        """Extract QuickStyleDef elements for style lookup."""
        styles = {}
        
        for style_def in root.findall('.//one:QuickStyleDef', self.namespace):
            style_info = {}
            known_style_attrs = {
                'index', 'name', 'fontColor', 'highlightColor', 'font', 
                'fontSize', 'spaceBefore', 'spaceAfter'
            }
            
            for attr, value in style_def.attrib.items():
                if attr in known_style_attrs:
                    style_info[attr] = value
                else:
                    self.unknown_attributes.add(f"QuickStyleDef.{attr}")
            
            index = style_def.attrib.get('index', '0')
            styles[index] = style_info
        
        return styles
    
    def _extract_title(self, root) -> Optional[str]:
        """Extract page title from Title element."""
        title_elem = root.find('.//one:Title/one:OE/one:T', self.namespace)
        if title_elem is not None and title_elem.text:
            return self._clean_cdata_content(title_elem.text)
        return None
    
    def _extract_content_hierarchy(self, root) -> List[Dict[str, Any]]:
        """Extract content preserving OE hierarchy."""
        content = []
        
        # Process all Outline elements
        for outline in root.findall('.//one:Outline', self.namespace):
            outline_content = self._process_outline_element(outline)
            if outline_content:
                content.extend(outline_content)
        
        return content
    
    def _process_outline_element(self, outline_elem) -> List[Dict[str, Any]]:
        """Process Outline element and its OEChildren hierarchy."""
        content = []
        
        # Extract outline metadata
        outline_meta = self._extract_outline_metadata(outline_elem)
        
        # Process OEChildren
        oe_children = outline_elem.find('one:OEChildren', self.namespace)
        if oe_children is not None:
            content.extend(self._process_oe_children(oe_children, level=0, outline_meta=outline_meta))
        
        return content
    
    def _extract_outline_metadata(self, outline_elem) -> Dict[str, Any]:
        """Extract metadata from Outline element."""
        meta = {}
        
        # Extract Position and Size elements
        position_elem = outline_elem.find('one:Position', self.namespace)
        if position_elem is not None:
            meta['position'] = position_elem.attrib
        
        size_elem = outline_elem.find('one:Size', self.namespace)
        if size_elem is not None:
            meta['size'] = size_elem.attrib
        
        # Extract outline attributes
        known_outline_attrs = {'author', 'authorInitials', 'lastModifiedBy', 'lastModifiedTime', 'objectID'}
        for attr, value in outline_elem.attrib.items():
            if attr in known_outline_attrs:
                if attr == 'lastModifiedTime':
                    try:
                        meta[attr] = self._parse_datetime(value)
                    except:
                        meta[attr] = value
                else:
                    meta[attr] = value
            else:
                self.unknown_attributes.add(f"Outline.{attr}")
        
        return meta
    
    def _process_oe_children(self, oe_children_elem, level: int = 0, outline_meta: Dict = None) -> List[Dict[str, Any]]:
        """Process OEChildren elements recursively."""
        content = []
        
        for oe in oe_children_elem.findall('one:OE', self.namespace):
            oe_content = self._process_oe_element(oe, level, outline_meta)
            if oe_content:
                content.append(oe_content)
        
        return content
    
    def _process_oe_element(self, oe_elem, level: int, outline_meta: Dict = None) -> Optional[Dict[str, Any]]:
        """Process a single OE (Outline Element) with full attribute handling."""
        oe_data = {
            'type': 'outline_element',
            'level': level,
            'attributes': {},
            'content': None,
            'children': []
        }
        
        # Extract all OE attributes
        known_oe_attrs = {
            'objectID', 'alignment', 'quickStyleIndex', 'creationTime', 'lastModifiedTime',
            'author', 'authorInitials', 'lastModifiedBy', 'lastModifiedByInitials',
            'authorResolutionID', 'lastModifiedByResolutionID', 'style', 'lang'
        }
        
        for attr, value in oe_elem.attrib.items():
            if attr in known_oe_attrs:
                if attr in ['creationTime', 'lastModifiedTime']:
                    try:
                        oe_data['attributes'][attr] = self._parse_datetime(value)
                    except:
                        oe_data['attributes'][attr] = value
                else:
                    oe_data['attributes'][attr] = value
            else:
                self.unknown_attributes.add(f"OE.{attr}")
        
        # Process child elements
        text_content = None
        has_image = False
        has_table = False
        
        for child in oe_elem:
            child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if child_tag == 'T':  # Text content
                if child.text and child.text != 'None':
                    cleaned_text = self._clean_cdata_content(child.text)
                    if cleaned_text.strip():  # Only use non-empty text
                        if text_content:
                            text_content += " " + cleaned_text  # Combine multiple T elements
                        else:
                            text_content = cleaned_text
            
            elif child_tag == 'Image':  # Image element
                has_image = True
                oe_data['content'] = self._process_image_element(child)
                oe_data['type'] = 'image'
            
            elif child_tag == 'Table':  # Handle Table elements
                has_table = True
                table_data = self._process_table_element(child)
                oe_data['content'] = table_data
                oe_data['type'] = 'table'
            
            elif child_tag == 'OEChildren':  # Nested children
                oe_data['children'] = self._process_oe_children(child, level + 1, outline_meta)
            
            else:
                # Log unknown child elements
                self.unknown_elements.add(f"OE.{child_tag}")
        
        # Set text content if found and no image/table
        if text_content and not has_image and not has_table:
            oe_data['content'] = text_content
            oe_data['type'] = 'text'
        
        # Only return if there's meaningful content
        if oe_data['content'] or oe_data['children'] or oe_data['attributes']:
            return oe_data
        
        return None
    
    def _process_image_element(self, image_elem) -> Dict[str, Any]:
        """Process Image element with CallbackID handling."""
        image_data = {
            'type': 'image',
            'alt': image_elem.attrib.get('alt', ''),
            'callback_id': None
        }
        
        # Find CallbackID
        callback_elem = image_elem.find('one:CallbackID', self.namespace)
        if callback_elem is not None:
            image_data['callback_id'] = callback_elem.attrib.get('callbackID')
        
        # Log unknown attributes
        known_img_attrs = {'alt'}
        for attr in image_elem.attrib:
            if attr not in known_img_attrs:
                self.unknown_attributes.add(f"Image.{attr}")
        
        return image_data
    
    def _process_table_element(self, table_elem) -> Dict[str, Any]:
        """Process Table element and extract all content including images."""
        table_data = {
            'type': 'table',
            'rows': [],
            'has_images': False,
            'embedded_images': []
        }
        
        # Process table attributes
        table_attrs = {}
        for attr, value in table_elem.attrib.items():
            table_attrs[attr] = value
        table_data['attributes'] = table_attrs
        
        # Process rows
        for row_elem in table_elem.findall('one:Row', self.namespace):
            row_data = []
            
            # Process cells in row
            for cell_elem in row_elem.findall('one:Cell', self.namespace):
                cell_content = self._process_cell_content(cell_elem)
                row_data.append(cell_content)
                
                # Check if cell contains images
                if isinstance(cell_content, dict) and cell_content.get('has_images'):
                    table_data['has_images'] = True
                    table_data['embedded_images'].extend(cell_content.get('images', []))
            
            table_data['rows'].append(row_data)
        
        return table_data
    
    def _process_cell_content(self, cell_elem) -> Dict[str, Any]:
        """Process cell content, extracting text and images."""
        cell_data = {
            'text': '',
            'images': [],
            'has_images': False
        }
        
        # Check for OEChildren in cell
        oe_children = cell_elem.find('one:OEChildren', self.namespace)
        if oe_children is not None:
            # Process each OE in the cell
            for oe in oe_children.findall('one:OE', self.namespace):
                # Check for text
                text_elem = oe.find('one:T', self.namespace)
                if text_elem is not None and text_elem.text:
                    cell_data['text'] += self._clean_cdata_content(text_elem.text) + ' '
                
                # Check for images
                image_elem = oe.find('one:Image', self.namespace)
                if image_elem is not None:
                    image_data = self._process_image_element(image_elem)
                    cell_data['images'].append(image_data)
                    cell_data['has_images'] = True
        
        return cell_data
    
    def _extract_images(self, root) -> List[Dict[str, Any]]:
        """Extract all image references from the page."""
        images = []
        
        for img_elem in root.findall('.//one:Image', self.namespace):
            image_data = self._process_image_element(img_elem)
            images.append(image_data)
        
        return images
    
    def _clean_cdata_content(self, text: str) -> str:
        """Clean CDATA content while preserving HTML for markdown conversion."""
        if not text:
            return ""
        
        # Normalize whitespace but preserve structure
        text = re.sub(r'\n\s*\n', '\n\n', text)
        text = re.sub(r' +', ' ', text)
        
        # Keep HTML tags intact for markdown conversion
        return text.strip()
    
    def _parse_datetime(self, dt_string: str) -> str:
        """Parse OneNote datetime format."""
        try:
            dt = datetime.fromisoformat(dt_string.replace('Z', '+00:00'))
            return dt.isoformat()
        except:
            return dt_string
    
    def get_unknown_patterns(self) -> Dict[str, List[str]]:
        """Get summary of unknown XML patterns discovered."""
        return {
            'unknown_elements': sorted(list(self.unknown_elements)),
            'unknown_attributes': sorted(list(self.unknown_attributes)),
            'discovery_timestamp': datetime.now().isoformat()
        }


def main():
    """Test the enhanced XML parser with exported files."""
    parser = OneNoteXMLParser()
    
    # Look for XML files in the Personal_XML directory
    output_dir = Path(__file__).parent.parent.parent / 'output' / 'Personal_XML'
    
    if not output_dir.exists():
        print(f"XML directory not found: {output_dir}")
        print("Run the PowerShell export script first: export_xml_only.ps1")
        return
    
    # Collect XML files from all sections
    xml_files = []
    for section_dir in output_dir.iterdir():
        if section_dir.is_dir():
            xml_files.extend(list(section_dir.glob('*.xml'))[:2])  # Limit for testing
    
    if not xml_files:
        print("No XML files found in Personal_XML directory")
        return
    
    print(f"Found {len(xml_files)} XML files to process")
    
    for xml_file in xml_files[:5]:  # Process max 5 files for testing
        print(f"\nParsing: {xml_file.relative_to(output_dir)}")
        print("=" * 60)
        
        try:
            result = parser.parse_page_xml(xml_file)
            
            print(f"Page: {result['page_name']}")
            print(f"Title: {result.get('title', 'No title')}")
            print(f"Content items: {result['parsing_stats']['content_items']}")
            print(f"Images: {result['parsing_stats']['image_count']}")
            print(f"Styles: {result['parsing_stats']['style_count']}")
            print(f"Unknown elements: {result['parsing_stats']['unknown_element_count']}")
            print(f"Unknown attributes: {result['parsing_stats']['unknown_attribute_count']}")
            
            # Show content preview
            if result['content']:
                print("\nContent preview:")
                for i, item in enumerate(result['content'][:3]):
                    content_text = str(item.get('content', ''))[:100]
                    print(f"  {i+1}. [{item['type']}] Level {item.get('level', 0)}: {content_text}...")
            
            # Show image info
            if result['images']:
                print(f"\nImages found:")
                for i, img in enumerate(result['images'][:3]):
                    callback_id = img.get('callback_id', 'No ID')[:20]
                    print(f"  {i+1}. {img.get('alt', 'No alt')} (CallbackID: {callback_id}...)")
            
            # Save parsed result
            output_path = xml_file.parent / f"{xml_file.stem}_parsed.json"
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2, ensure_ascii=False, default=str)
            
            print(f"\nParsed content saved to: {output_path}")
            
        except Exception as e:
            print(f"Error parsing {xml_file}: {e}")
            import traceback
            traceback.print_exc()
    
    # Print discovery summary
    print("\n" + "=" * 60)
    print("UNKNOWN PATTERN DISCOVERY SUMMARY")
    print("=" * 60)
    
    patterns = parser.get_unknown_patterns()
    
    if patterns['unknown_elements']:
        print(f"\nUnknown XML elements ({len(patterns['unknown_elements'])}):")
        for elem in patterns['unknown_elements']:
            print(f"  - {elem}")
    
    if patterns['unknown_attributes']:
        print(f"\nUnknown attributes ({len(patterns['unknown_attributes'])}):")
        for attr in patterns['unknown_attributes']:
            print(f"  - {attr}")
    
    if not patterns['unknown_elements'] and not patterns['unknown_attributes']:
        print("\nAll XML patterns recognized - parser is comprehensive!")
    
    print(f"\nDiscovery completed: {patterns['discovery_timestamp']}")


if __name__ == "__main__":
    main()