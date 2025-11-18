#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNote Extractors Module.

Provides extractors for parsing OneNote XML exports and
extracting binary content (images) via the OneNote COM API.

Available extractors:
    - OneNoteXMLParser: Parse OneNote XML page exports
    - OneNoteImageExtractor: Extract images using CallbackIDs
"""

__all__ = [
    "OneNoteXMLParser",
    "OneNoteImageExtractor",
]

# Optional: Import main classes for easier access
# Uncomment if you want to enable: from extractors import OneNoteXMLParser
# from .onenote_xml_parser import OneNoteXMLParser
# from .image_extractor import OneNoteImageExtractor
