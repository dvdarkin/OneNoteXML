#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNote Image Extractor - Extract actual images using OneNote COM API
Uses CallbackIDs from XML to get binary image data
"""

import base64
import hashlib
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

class OneNoteImageExtractor:
    """Extract images from OneNote using COM API and CallbackIDs."""
    
    def __init__(self):
        self.onenote = None
        self.logger = logging.getLogger('ImageExtractor')
        
    def __enter__(self):
        """Context manager entry - initialize OneNote COM object."""
        if not COM_AVAILABLE:
            raise RuntimeError("pywin32 not available - cannot use OneNote COM API")
        
        try:
            self.onenote = win32com.client.Dispatch("OneNote.Application")
            self.logger.info("OneNote COM interface initialized")
            return self
        except Exception as e:
            raise RuntimeError(f"Failed to initialize OneNote COM: {e}")
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleanup COM object."""
        if self.onenote:
            try:
                # Release COM object
                del self.onenote
                self.onenote = None
                self.logger.info("OneNote COM interface released")
            except:
                pass
    
    def extract_images_from_page(self, page_id: str, callback_ids: List[str], 
                                assets_dir: Path) -> Dict[str, str]:
        """
        Extract images from a page using CallbackIDs.
        
        Args:
            page_id: OneNote page ID
            callback_ids: List of CallbackIDs to extract
            assets_dir: Directory to save extracted images
            
        Returns:
            Dict mapping CallbackID to saved filename
        """
        if not self.onenote:
            raise RuntimeError("OneNote COM interface not initialized")
        
        assets_dir.mkdir(parents=True, exist_ok=True)
        extracted_images = {}
        
        for callback_id in callback_ids:
            try:
                image_filename = self._extract_single_image(page_id, callback_id, assets_dir)
                if image_filename:
                    extracted_images[callback_id] = image_filename
                    self.logger.info(f"Extracted image: {callback_id} -> {image_filename}")
                else:
                    self.logger.warning(f"Failed to extract image: {callback_id}")
                    
            except Exception as e:
                self.logger.error(f"Error extracting image {callback_id}: {e}")
                # Continue with other images
        
        return extracted_images
    
    def _extract_single_image(self, page_id: str, callback_id: str, 
                             assets_dir: Path) -> Optional[str]:
        """Extract a single image and save to assets directory."""
        try:
            # Get binary content from OneNote
            binary_data_b64 = self.onenote.GetBinaryPageContent(page_id, callback_id)
            
            if not binary_data_b64:
                self.logger.warning(f"No binary data returned for {callback_id}")
                return None
            
            # Decode base64 data
            try:
                image_data = base64.b64decode(binary_data_b64)
            except Exception as e:
                self.logger.error(f"Failed to decode base64 data for {callback_id}: {e}")
                return None
            
            if len(image_data) == 0:
                self.logger.warning(f"Empty image data for {callback_id}")
                return None
            
            # Determine image format from header
            image_format = self._detect_image_format(image_data)
            
            # Generate filename
            # Use content hash to avoid duplicates and ensure consistent naming
            content_hash = hashlib.md5(image_data).hexdigest()[:12]
            safe_callback_id = re.sub(r'[^\w-]', '_', callback_id)[:20]
            filename = f"{safe_callback_id}_{content_hash}.{image_format}"
            
            # Save image file
            image_path = assets_dir / filename
            with open(image_path, 'wb') as f:
                f.write(image_data)
            
            # Log success with file size
            file_size = len(image_data)
            self.logger.debug(f"Saved {filename}: {file_size} bytes, format: {image_format}")
            
            return filename
            
        except Exception as e:
            self.logger.error(f"Error in _extract_single_image for {callback_id}: {e}")
            return None
    
    def _detect_image_format(self, image_data: bytes) -> str:
        """Detect image format from binary data headers."""
        if len(image_data) < 8:
            return "bin"  # Unknown format
        
        # Check common image format signatures
        if image_data.startswith(b'\xFF\xD8\xFF'):
            return "jpg"
        elif image_data.startswith(b'\x89PNG\r\n\x1A\n'):
            return "png"
        elif image_data.startswith(b'GIF87a') or image_data.startswith(b'GIF89a'):
            return "gif"
        elif image_data.startswith(b'BM'):
            return "bmp"
        elif image_data.startswith(b'\x00\x00\x01\x00'):
            return "ico"
        elif image_data.startswith(b'RIFF') and b'WEBP' in image_data[:12]:
            return "webp"
        else:
            # Try to detect by common patterns
            if b'JFIF' in image_data[:20] or b'Exif' in image_data[:20]:
                return "jpg"
            elif b'PNG' in image_data[:20]:
                return "png"
            else:
                return "bin"  # Unknown format
    
    def get_page_images_info(self, parsed_page_data: Dict) -> List[Dict]:
        """
        Extract image information from parsed page data.
        
        Args:
            parsed_page_data: Parsed page data from XML parser
            
        Returns:
            List of dicts with image info (callback_id, alt_text, etc.)
        """
        images_info = []
        
        # Get page ID for COM API calls
        page_id = parsed_page_data.get('metadata', {}).get('ID')
        if not page_id:
            self.logger.warning("No page ID found in metadata")
            return images_info
        
        # Extract images from the images list
        for image in parsed_page_data.get('images', []):
            callback_id = image.get('callback_id')
            if callback_id:
                images_info.append({
                    'callback_id': callback_id,
                    'alt_text': image.get('alt', 'Image'),
                    'page_id': page_id
                })
        
        # Also check content items for inline images
        for item in parsed_page_data.get('content', []):
            if item.get('type') == 'image' and isinstance(item.get('content'), dict):
                callback_id = item['content'].get('callback_id')
                if callback_id:
                    images_info.append({
                        'callback_id': callback_id,
                        'alt_text': item['content'].get('alt', 'Image'),
                        'page_id': page_id
                    })
        
        return images_info


def extract_images_for_page(parsed_page_data: Dict, assets_dir: Path) -> Dict[str, str]:
    """
    Convenience function to extract all images for a page.
    
    Args:
        parsed_page_data: Parsed page data from XML parser
        assets_dir: Directory to save images
        
    Returns:
        Dict mapping CallbackID to saved filename
    """
    if not COM_AVAILABLE:
        logging.getLogger('ImageExtractor').warning(
            "pywin32 not available - cannot extract images"
        )
        return {}
    
    try:
        with OneNoteImageExtractor() as extractor:
            images_info = extractor.get_page_images_info(parsed_page_data)
            
            if not images_info:
                return {}
            
            # Extract all CallbackIDs
            callback_ids = [img['callback_id'] for img in images_info]
            page_id = images_info[0]['page_id']  # All images are from same page
            
            return extractor.extract_images_from_page(page_id, callback_ids, assets_dir)
            
    except Exception as e:
        logging.getLogger('ImageExtractor').error(f"Failed to extract images: {e}")
        return {}


def main():
    """Test the image extractor with sample data."""
    logging.basicConfig(level=logging.INFO)
    
    print("OneNote Image Extractor Test")
    print("=" * 40)
    
    if not COM_AVAILABLE:
        print("ERROR: pywin32 not available")
        print("Install with: pip install pywin32")
        return
    
    # Test with a sample page that has images
    import sys
    sys.path.append('../extractors')
    from onenote_xml_parser import OneNoteXMLParser
    
    parser = OneNoteXMLParser()
    
    # Look for XML files with images
    xml_dir = Path(__file__).parent.parent.parent / 'output' / 'Personal_XML'
    
    for xml_file in xml_dir.rglob('*.xml'):
        try:
            parsed_data = parser.parse_page_xml(xml_file)
            if parsed_data.get('images'):
                print(f"\nTesting with: {xml_file.name}")
                print(f"Found {len(parsed_data['images'])} images")
                
                # Create test assets directory
                test_assets = Path('test_assets')
                
                # Extract images
                extracted = extract_images_for_page(parsed_data, test_assets)
                
                print(f"Successfully extracted {len(extracted)} images:")
                for callback_id, filename in extracted.items():
                    print(f"  {callback_id[:20]}... -> {filename}")
                
                break  # Test with first page that has images
                
        except Exception as e:
            print(f"Error testing {xml_file.name}: {e}")
    
    else:
        print("No pages with images found for testing")


if __name__ == "__main__":
    main()