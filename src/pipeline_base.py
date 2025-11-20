#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
Shared utilities for OneNote to Markdown conversion pipelines.
Common functions used by both Obsidian and Logseq pipelines.
"""

import sys
import logging
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Tuple


def setup_logging(output_dir: Path, logger_name: str) -> logging.Logger:
    """
    Set up logging configuration with UTF-8 encoding.

    Args:
        output_dir: Base output directory for logs
        logger_name: Name for the logger (e.g., 'OneNoteObsidian')

    Returns:
        Configured logger instance
    """
    log_dir = output_dir / 'logs'
    log_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = log_dir / f'{logger_name.lower()}_{timestamp}.log'

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    # Set console output encoding to UTF-8 for Windows
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')

    return logging.getLogger(logger_name)


def parse_pipeline_args(script_name: str) -> Tuple[str, Path]:
    """
    Parse command line arguments for pipeline scripts.

    Args:
        script_name: Name of the script for usage message

    Returns:
        Tuple of (notebook_name, output_base_dir)

    Raises:
        SystemExit: If arguments are invalid
    """
    if len(sys.argv) < 3:
        print(f"Usage: python {script_name} <notebook_name> <output_dir>")
        print(f"Example: python {script_name} 'Personal' 'output/Personal'")
        sys.exit(1)

    notebook_name = sys.argv[1]
    output_base_dir = Path(sys.argv[2])

    return notebook_name, output_base_dir


def group_pages_by_section(xml_files: List[Path]) -> Dict[str, List[Path]]:
    """
    Group XML files by their section (parent directory).
    Sort pages by numeric prefix to preserve OneNote hierarchy order.

    Args:
        xml_files: List of XML file paths

    Returns:
        Dictionary mapping section names to lists of XML files (sorted by numeric prefix)
    """
    sections = defaultdict(list)

    for xml_file in xml_files:
        section_name = xml_file.parent.name
        sections[section_name].append(xml_file)

    # Sort each section's pages by numeric prefix (001_, 002_, etc.)
    for section_name in sections:
        sections[section_name] = sort_pages_by_hierarchy(sections[section_name])

    return dict(sections)


def sort_pages_by_hierarchy(xml_files: List[Path]) -> List[Path]:
    """
    Sort XML files by numeric prefix to maintain OneNote hierarchy order.

    Files with numeric prefix (e.g., '001_Main.xml') sort by number.
    Files without prefix sort alphabetically at the end.

    Args:
        xml_files: List of XML file paths

    Returns:
        Sorted list of XML files
    """
    import re

    def get_sort_key(path: Path) -> tuple:
        # Extract numeric prefix if present (e.g., "001" from "001_Main.xml")
        match = re.match(r'^(\d+)_', path.name)
        if match:
            return (0, int(match.group(1)))  # (has_prefix, number)
        else:
            return (1, path.name)  # (no_prefix, alphabetical)

    return sorted(xml_files, key=get_sort_key)


def discover_xml_files(xml_input_dir: Path, logger: logging.Logger) -> List[Path]:
    """
    Discover all XML files in section subdirectories.

    Args:
        xml_input_dir: Directory containing XML section folders
        logger: Logger instance for reporting

    Returns:
        List of XML file paths

    Raises:
        SystemExit: If directory doesn't exist or no files found
    """
    # Check if XML directory exists
    if not xml_input_dir.exists():
        logger.error(f"XML input directory not found: {xml_input_dir}")
        print(f"\nError: XML directory not found at {xml_input_dir}")
        print("The XML export step may have failed.")
        sys.exit(1)

    # Find XML files in all section subdirectories
    xml_files = []
    for section_dir in xml_input_dir.iterdir():
        if section_dir.is_dir():
            xml_files.extend(list(section_dir.glob('*.xml')))

    # Validate files were found
    if not xml_files:
        logger.warning(f"No XML files found in {xml_input_dir}")
        print(f"\nNo XML files found in {xml_input_dir}")
        print("The XML export step may have failed.")
        sys.exit(1)

    logger.info(f"Found {len(xml_files)} XML file(s) to process")

    return xml_files


def log_pipeline_start(logger: logging.Logger, pipeline_name: str,
                       notebook_name: str, output_dir: Path):
    """
    Log the start of a pipeline execution.

    Args:
        logger: Logger instance
        pipeline_name: Name of the pipeline (e.g., "Obsidian Vault Generator")
        notebook_name: Name of the notebook being processed
        output_dir: Output directory path
    """
    logger.info(f"OneNoteXML - {pipeline_name}")
    logger.info("=" * 70)
    logger.info(f"Notebook: {notebook_name}")
    logger.info(f"Output Directory: {output_dir}")


def log_conversion_summary(logger: logging.Logger, total_success: int,
                           total_files: int, sections: Dict):
    """
    Log the final conversion summary.

    Args:
        logger: Logger instance
        total_success: Number of successfully processed files
        total_files: Total number of files
        sections: Dictionary of sections processed
    """
    logger.info("=" * 70)
    logger.info(f"Processing complete: {total_success}/{total_files} files processed successfully")

    if total_success > 0:
        print(f"\nSuccessfully processed {total_success} page(s)")
        print(f"Converted {len(sections)} section(s)")
    else:
        print("\nNo files processed successfully")
