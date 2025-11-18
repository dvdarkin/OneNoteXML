#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNoteXML - Direct XML extraction from OneNote to Markdown.

This package provides tools for extracting OneNote notebooks to
Obsidian and Logseq markdown formats using direct XML extraction
via the OneNote COM API.

Modules:
    - extractors: XML parsing and image extraction
    - converters: Markdown conversion (Obsidian and Logseq)
    - pipeline_base: Shared utilities for conversion pipelines
    - obsidian_pipeline: Obsidian vault generation
    - logseq_pipeline: Logseq graph generation
"""

__version__ = "1.0.0"
__author__ = "Denis Darkin"
__license__ = "MIT"

__all__ = [
    "extractors",
    "converters",
    "pipeline_base",
    "obsidian_pipeline",
    "logseq_pipeline",
]
