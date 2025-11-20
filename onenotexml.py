#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
OneNoteXML - Direct XML extraction from OneNote to Markdown

Usage:
    python onenotexml.py NOTEBOOK_NAME [options]

Examples:
    python onenotexml.py "Personal"
    python onenotexml.py "Project1" --format logseq
    python onenotexml.py "Research" --output ./my-vault --format obsidian

Why OneNoteXML?
    - Direct XML extraction, flexible
    - Working image extraction (CallbackID-based)
    - Dual output: Obsidian OR Logseq

Requirements:
    - Windows OS (OneNote COM API)
    - OneNote 2010-2013 desktop
    - Python 3.8+
"""

import sys
import argparse
import subprocess
from pathlib import Path
import logging

def check_platform():
    """Verify Windows platform."""
    if sys.platform != 'win32':
        print("ERROR: Windows required (OneNote COM API dependency)")
        print("   This tool uses OneNote's COM interface which only works on Windows.")
        return False
    return True

def check_python_version():
    """Verify Python 3.8+."""
    if sys.version_info < (3, 8):
        print(f"ERROR: Python 3.8+ required (you have {sys.version_info.major}.{sys.version_info.minor})")
        return False
    return True

def check_onenote():
    """Verify OneNote is accessible via COM."""
    try:
        import win32com.client
        onenote = win32com.client.Dispatch("OneNote.Application")
        del onenote
        print("OK: OneNote COM API accessible")
        return True
    except ImportError:
        print("ERROR: pywin32 not installed")
        print("   Install: pip install pywin32")
        return False
    except Exception as e:
        print("ERROR: Cannot access OneNote COM interface")
        print("   Is OneNote 2010-2013 desktop version installed?")
        print(f"   Details: {e}")
        return False

def check_requirements():
    """Run all requirement checks."""
    print("OneNoteXML - Checking requirements...")
    print("=" * 60)

    checks = [
        ("Platform", check_platform()),
        ("Python version", check_python_version()),
        ("OneNote COM", check_onenote()),
    ]

    all_passed = all(result for _, result in checks)

    if all_passed:
        print("=" * 60)
        print("All requirements met\n")
    else:
        print("=" * 60)
        print("Requirements not met. Please fix errors above.\n")

    return all_passed

def run_subprocess_with_progress(cmd, timeout=300, show_progress=True):
    """Run subprocess and stream output in real-time.

    Args:
        cmd: Command list to execute
        timeout: Timeout in seconds
        show_progress: If True, print output as it arrives

    Returns:
        tuple: (returncode, stdout, stderr)
    """
    import threading

    # Start process
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    # Queues to collect output from threads
    stdout_lines = []
    stderr_lines = []

    def read_stream(stream, output_list, prefix=""):
        """Read stream line by line and collect."""
        try:
            for line in stream:
                line = line.rstrip()
                if line:  # Skip empty lines
                    output_list.append(line)
                    if show_progress:
                        # Indent subprocess output for clarity
                        print(f"      {line}")
        except Exception:
            pass  # Stream closed or encoding error

    # Start reader threads
    stdout_thread = threading.Thread(
        target=read_stream,
        args=(process.stdout, stdout_lines)
    )
    stderr_thread = threading.Thread(
        target=read_stream,
        args=(process.stderr, stderr_lines)
    )

    stdout_thread.daemon = True
    stderr_thread.daemon = True
    stdout_thread.start()
    stderr_thread.start()

    # Wait for process with timeout
    try:
        returncode = process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait()
        raise

    # Wait for reader threads to finish
    stdout_thread.join(timeout=1)
    stderr_thread.join(timeout=1)

    return returncode, '\n'.join(stdout_lines), '\n'.join(stderr_lines)

def run_extraction(notebook_name: str, output_format: str, output_dir: Path, debug: bool = False):
    """Run the complete extraction pipeline.

    Args:
        notebook_name: Name of the OneNote notebook
        output_format: Output format ('obsidian' or 'logseq')
        output_dir: Base output directory
        debug: If True, keep interim files (XML, images) in debug folder
    """

    print(f"\nOneNoteXML - Extracting '{notebook_name}'")
    print("=" * 60)

    # Determine output paths
    notebook_output = output_dir / notebook_name
    xml_dir = notebook_output / "XML"
    images_dir = notebook_output / "images"
    vault_dir = notebook_output / f"{output_format}_vault"

    # Create directories
    for dir_path in [notebook_output, xml_dir, images_dir, vault_dir]:
        dir_path.mkdir(parents=True, exist_ok=True)

    # Step 1: Export XML from OneNote
    print("\n[1/3] Exporting XML from OneNote...")
    print(f"      → {xml_dir}")

    # Validate PowerShell script exists
    ps_script = Path(__file__).parent / "scripts" / "export_xml_notebook.ps1"
    if not ps_script.exists():
        print(f"ERROR: PowerShell script not found: {ps_script}")
        print(f"   Expected location: {ps_script.absolute()}")
        return False

    # Verify output directory is writable
    try:
        xml_dir.mkdir(parents=True, exist_ok=True)
        test_file = xml_dir / ".write_test"
        test_file.touch()
        test_file.unlink()
    except (IOError, OSError, PermissionError) as e:
        print(f"ERROR: Cannot write to output directory: {xml_dir}")
        print(f"   {e}")
        return False

    try:
        returncode, stdout, stderr = run_subprocess_with_progress(
            ["PowerShell", "-ExecutionPolicy", "Bypass", "-File",
             str(ps_script), "-NotebookName", notebook_name,
             "-OutputPath", str(xml_dir)],
            timeout=300,
            show_progress=True
        )

        if returncode != 0:
            print(f"ERROR: XML export failed (exit code: {returncode})")
            if stderr:
                print(f"   Error details:")
                for line in stderr.strip().split('\n')[:10]:  # Show first 10 lines
                    print(f"     {line}")
            return False

        print("      XML export completed")

    except subprocess.TimeoutExpired:
        print("ERROR: XML export timed out (>5 minutes)")
        return False
    except Exception as e:
        print(f"ERROR: {e}")
        return False

    # Step 2: Convert to markdown (Obsidian or Logseq)
    print(f"\n[2/3] Converting to {output_format} format...")
    print(f"      → {vault_dir}")

    if output_format == "obsidian":
        converter_script = Path(__file__).parent / "src" / "obsidian_pipeline.py"
    else:  # logseq
        converter_script = Path(__file__).parent / "src" / "logseq_pipeline.py"

    if not converter_script.exists():
        print(f"ERROR: Converter script not found: {converter_script}")
        return False

    try:
        returncode, stdout, stderr = run_subprocess_with_progress(
            ["python", str(converter_script), notebook_name, str(notebook_output)],
            timeout=300,
            show_progress=True
        )

        if returncode != 0:
            print(f"ERROR: Conversion failed (exit code: {returncode})")
            if stderr:
                print(f"   Error details:")
                for line in stderr.strip().split('\n')[:10]:
                    print(f"     {line}")
            return False

        print(f"      {output_format.title()} conversion completed")

    except Exception as e:
        print(f"ERROR: {e}")
        return False

    # Step 3: Extract images
    print("\n[3/3] Extracting images...")
    print(f"      → {images_dir}")

    image_map = vault_dir / "image_extraction_map.json"
    if not image_map.exists():
        print("      No images to extract")
        return True

    ps_image_script = Path(__file__).parent / "scripts" / "extract_images_robust.ps1"
    if not ps_image_script.exists():
        print(f"WARNING: Image extraction script not found: {ps_image_script}")
        print("      Images will not be extracted")
        return True

    try:
        returncode, stdout, stderr = run_subprocess_with_progress(
            ["PowerShell", "-ExecutionPolicy", "Bypass", "-File",
             str(ps_image_script), "-NotebookName", notebook_name,
             "-OutputPath", str(images_dir), "-MapFile", str(image_map)],
            timeout=600,  # 10 minute timeout for images
            show_progress=True
        )

        if returncode != 0:
            print(f"WARNING: Some images may not have extracted")
            print(f"   This is normal - OneNote stores some images externally")
        else:
            print("      Image extraction completed")

        # Copy images to vault (do this regardless of extraction result)
        if output_format == "obsidian":
            attachments_dir = vault_dir / f"{notebook_name}-Vault" / "attachments"
            try:
                attachments_dir.mkdir(parents=True, exist_ok=True)

                import shutil
                import os

                # Get list of images to copy
                images_to_copy = list(images_dir.glob("*.*"))
                copied_count = 0
                failed_count = 0

                if images_to_copy:
                    for img in images_to_copy:
                        try:
                            # Verify source file exists and is readable
                            if not img.is_file():
                                continue

                            # Copy with error handling
                            dest_path = attachments_dir / img.name
                            shutil.copy2(img, dest_path)

                            # Verify copy succeeded
                            if dest_path.exists() and dest_path.stat().st_size > 0:
                                copied_count += 1
                            else:
                                failed_count += 1
                                print(f"      Failed to copy: {img.name}")
                        except (IOError, OSError) as e:
                            failed_count += 1
                            print(f"      Error copying {img.name}: {e}")

                    # Report results
                    print(f"      Images copied: {copied_count} of {len(images_to_copy)}")
                    if failed_count > 0:
                        print(f"      Failed to copy: {failed_count} images")
                else:
                    print(f"      No images found in {images_dir}")

            except Exception as e:
                print(f"      Error setting up attachments: {e}")

    except Exception as e:
        print(f"WARNING: Image extraction issue: {e}")
        print("   Continuing anyway...")

    # Final verification
    print(f"\n[Verification]")
    try:
        # Count markdown files
        if output_format == "obsidian":
            vault_path = vault_dir / f"{notebook_name}-Vault"
            md_files = list(vault_path.rglob("*.md"))
            attachments = list((vault_path / "attachments").glob("*.*")) if (vault_path / "attachments").exists() else []
        else:  # logseq
            md_files = list((vault_dir / "pages").rglob("*.md")) if (vault_dir / "pages").exists() else []
            attachments = list((vault_dir / "assets").glob("*.*")) if (vault_dir / "assets").exists() else []

        print(f"      Created {len(md_files)} markdown files")
        print(f"      Copied {len(attachments)} images to vault")

        if len(md_files) == 0:
            print(f"      WARNING: No markdown files found - extraction may have failed")
            return False

    except Exception as e:
        print(f"      Could not verify output: {e}")

    # Reorganize output structure
    print(f"\n[Cleanup & Organization]")
    try:
        import shutil

        # Define final paths
        if output_format == "obsidian":
            source_vault = vault_dir / f"{notebook_name}-Vault"
        else:  # logseq
            source_vault = vault_dir

        final_vault_path = output_dir / f"{notebook_name}-Vault"

        # Move vault to top level
        if source_vault.exists():
            if final_vault_path.exists():
                shutil.rmtree(final_vault_path)
            shutil.move(str(source_vault), str(final_vault_path))
            print(f"      Moved vault to: {final_vault_path.name}")

        # Handle interim files based on debug flag
        if debug:
            # Create debug directory and move interim files
            debug_dir = output_dir / f"{notebook_name}-debug"
            debug_dir.mkdir(parents=True, exist_ok=True)

            # Move XML files
            if xml_dir.exists():
                debug_xml = debug_dir / "XML"
                if debug_xml.exists():
                    shutil.rmtree(debug_xml)
                shutil.move(str(xml_dir), str(debug_xml))
                print(f"      Moved XML to debug folder")

            # Move staging images
            if images_dir.exists():
                debug_images = debug_dir / "images"
                if debug_images.exists():
                    shutil.rmtree(debug_images)
                shutil.move(str(images_dir), str(debug_images))
                print(f"      Moved staging images to debug folder")

            # Move image extraction map
            image_map = vault_dir / "image_extraction_map.json"
            if image_map.exists():
                shutil.move(str(image_map), str(debug_dir / "image_extraction_map.json"))
                print(f"      Moved metadata to debug folder")

            print(f"      Debug files saved to: {debug_dir.name}")

        else:
            # Delete interim files
            if xml_dir.exists():
                shutil.rmtree(xml_dir)
                print(f"      Deleted XML files ({16}MB)")

            if images_dir.exists():
                shutil.rmtree(images_dir)
                print(f"      Deleted staging images")

            image_map = vault_dir / "image_extraction_map.json"
            if image_map.exists():
                image_map.unlink()
                print(f"      Deleted metadata file")

        # Clean up empty directories
        if vault_dir.exists() and not any(vault_dir.iterdir()):
            vault_dir.rmdir()

        if notebook_output.exists() and not any(notebook_output.iterdir()):
            notebook_output.rmdir()
            print(f"      Cleaned up temporary directories")

    except Exception as e:
        print(f"      Warning during cleanup: {e}")
        print(f"      Vault is still available at: {vault_dir}")

    return True

def main():
    """Main entry point."""

    # Set console output encoding to UTF-8 for Windows
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except (AttributeError, OSError):
            # Fallback for older Python versions or if reconfigure fails
            pass

    parser = argparse.ArgumentParser(
        description="OneNoteXML - Extract OneNote notebooks to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python onenotexml.py "Personal"
    python onenotexml.py "Work Notes" --format logseq
    python onenotexml.py "Research" --output ./my-notes
    python onenotexml.py "Personal" --debug

About OneNoteXML:
    Direct XML extraction from OneNote to Obsidian/Logseq markdown.
    
    - Working image extraction (CallbackID-based)
    - Dual format support (Obsidian OR Logseq)
    - Minimal dependencies (BeautifulSoup + pywin32)
    - Local-only (no cloud sync required)

For more info: https://github.com/dvdarkin/OneNoteXML
        """
    )

    parser.add_argument(
        'notebook',
        help='OneNote notebook name (case-sensitive)'
    )

    parser.add_argument(
        '--format',
        choices=['obsidian', 'logseq'],
        default='obsidian',
        help='Output format (default: obsidian)'
    )

    parser.add_argument(
        '--output',
        type=Path,
        default=Path('./output'),
        help='Output directory (default: ./output)'
    )

    parser.add_argument(
        '--check-only',
        action='store_true',
        help='Only check requirements, do not extract'
    )

    parser.add_argument(
        '--debug',
        action='store_true',
        help='Keep interim files (XML, staging images) for debugging'
    )

    args = parser.parse_args()

    # Check requirements
    if not check_requirements():
        return 1

    if args.check_only:
        print("Requirements check passed. Ready to extract.")
        return 0

    # Run extraction
    print(f"\nNotebook: {args.notebook}")
    print(f"Format:   {args.format}")
    print(f"Output:   {args.output}")
    if args.debug:
        print(f"Debug:    enabled (keeping interim files)")

    success = run_extraction(args.notebook, args.format, args.output, args.debug)

    if success:
        print("\n" + "=" * 60)
        print("SUCCESS: Extraction complete!")
        print("=" * 60)

        vault_path = args.output / f"{args.notebook}-Vault"
        print(f"\nYour {args.format} vault is ready:")
        print(f"  {vault_path}")

        if args.debug:
            debug_path = args.output / f"{args.notebook}-debug"
            print(f"\nDebug files saved to:")
            print(f"  {debug_path}")

        logs_path = args.output / "logs"
        print(f"\nLogs saved to:")
        print(f"  {logs_path}")

        print(f"\nTo open in {args.format.title()}:")
        if args.format == "obsidian":
            print(f"  1. Open Obsidian")
            print(f"  2. Click 'Open folder as vault'")
            print(f"  3. Select: {vault_path}")
        else:
            print(f"  1. Open Logseq")
            print(f"  2. Add graph")
            print(f"  3. Select: {vault_path}")

        return 0
    else:
        print("\n" + "=" * 60)
        print("FAILED: Check errors above")
        print("=" * 60)
        print("\nCommon issues:")
        print("  • OneNote notebook not found (check spelling)")
        print("  • OneNote not running or not accessible")
        print("  • Notebook not downloaded locally (web-only notebooks won't work)")
        return 1

if __name__ == '__main__':
    sys.exit(main())
