#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNote Converters Module.

Provides converters for transforming parsed OneNote XML content
into various markdown formats.

Available converters:
    - MarkdownConverter: Base converter class
    - ObsidianConverter: Obsidian-compatible vault generation
    - LogseqConverter: Logseq-compatible graph generation
    - markdown_utils: Shared markdown sanitization utilities
"""

__all__ = [
    "MarkdownConverter",
    "ObsidianConverter",
    "LogseqConverter",
]

# Optional: Import main classes for easier access
# Uncomment if you want to enable: from converters import ObsidianConverter
# from .markdown_converter import MarkdownConverter
# from .obsidian_converter import ObsidianConverter
# from .logseq_converter import LogseqConverter
