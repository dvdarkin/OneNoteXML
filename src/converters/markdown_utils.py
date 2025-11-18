#!/usr/bin/env python3
# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT
"""
Shared utilities for markdown conversion and sanitization.
"""

import re
import html
from typing import List, Tuple


def _extract_code_blocks(text: str) -> List[Tuple[int, int, str]]:
    """Extract positions of code blocks and inline code.

    Returns:
        List of (start, end, type) tuples where type is 'fence' or 'inline'
    """
    code_regions = []

    # Find fenced code blocks (```...```)
    for match in re.finditer(r'```[\s\S]*?```', text):
        code_regions.append((match.start(), match.end(), 'fence'))

    # Find inline code (`...`)
    for match in re.finditer(r'`[^`\n]+?`', text):
        code_regions.append((match.start(), match.end(), 'inline'))

    # Sort by start position
    code_regions.sort(key=lambda x: x[0])
    return code_regions


def _is_in_code_context(position: int, code_regions: List[Tuple[int, int, str]]) -> bool:
    """Check if a position is inside a code block or inline code."""
    for start, end, _ in code_regions:
        if start <= position < end:
            return True
    return False


def escape_logseq_special_syntax(text: str) -> str:
    """Escape Logseq-specific syntax in non-code contexts.

    Escapes special Logseq syntax that could cause false triggers:
    - {{...}} (queries/commands)
    - word::value (properties)
    - ((block-ref)) (block references)

    Preserves these patterns inside:
    - Fenced code blocks (```...```)
    - Inline code (`...`)
    - URLs (http://, https://)

    Args:
        text: Text to sanitize

    Returns:
        Text with Logseq syntax escaped in non-code contexts

    Examples:
        >>> escape_logseq_special_syntax("{{title}} is a query")
        '\\\\{\\\\{title\\\\}\\\\} is a query'

        >>> escape_logseq_special_syntax("Use `{{var}}` in code")
        'Use `{{var}}` in code'  # Preserved in inline code

        >>> escape_logseq_special_syntax("head :: tail operator")
        'head \\\\:\\\\: tail operator'
    """
    # Extract code regions to preserve them
    code_regions = _extract_code_blocks(text)

    # Process each pattern
    result = text

    # 1. Escape double curly braces {{...}}
    # Find all {{ and }} outside code contexts
    for match in reversed(list(re.finditer(r'\{\{|\}\}', result))):
        if not _is_in_code_context(match.start(), code_regions):
            # Escape by doubling the backslashes
            escaped = '\\{\\{' if match.group() == '{{' else '\\}\\}'
            result = result[:match.start()] + escaped + result[match.end():]

    # 2. Escape double colons :: (but not in URLs or code)
    # Pattern: :: not preceded by : and not followed by /
    for match in reversed(list(re.finditer(r'(?<!:)::(?!/)', result))):
        if not _is_in_code_context(match.start(), code_regions):
            # Check if it's not part of a URL
            before = result[max(0, match.start()-10):match.start()]
            if 'http' not in before and 'https' not in before:
                result = result[:match.start()] + '\\:\\:' + result[match.end():]

    # 3. Escape double parentheses ((block-ref))
    for match in reversed(list(re.finditer(r'\(\(|\)\)', result))):
        if not _is_in_code_context(match.start(), code_regions):
            escaped = '\\(\\(' if match.group() == '((' else '\\)\\)'
            result = result[:match.start()] + escaped + result[match.end():]

    return result


def escape_literal_brackets_with_links(text: str) -> str:
    """Escape literal brackets that contain markdown links.

    Prevents syntax conflicts when content like [item1, item2, item3]
    contains markdown links [text](url) inside the brackets.

    This is a common issue when converting from OneNote where users write
    lists like [Link1, Link2, Link3] and each item is a hyperlink. When
    converted to markdown, this becomes [[Link1](url), [Link2](url), [Link3](url)]
    which creates confusing nested bracket syntax.

    Args:
        text: Text potentially containing brackets with markdown links

    Returns:
        Text with outer brackets escaped when they contain multiple links

    Examples:
        >>> escape_literal_brackets_with_links("[[List](url1), [Dict](url2)]")
        '\\\\[[List](url1), [Dict](url2)\\\\]'

        >>> escape_literal_brackets_with_links("See [[Documentation](url)]")
        'See [[Documentation](url)]'

        >>> escape_literal_brackets_with_links("data = [[x](url1), [y](url2)]")
        'data = \\\\[[x](url1), [y](url2)\\\\]'
    """
    # Pattern: Match brackets containing markdown links
    # Uses lazy matching to handle nested content
    # (?:[^\[\]]|\[[^\]]+\]\([^\)]+\))+ matches either:
    #   - Non-bracket characters [^\[\]]
    #   - OR markdown links \[[^\]]+\]\([^\)]+\)
    pattern = r'\[((?:[^\[\]]|\[[^\]]+\]\([^\)]+\))+)\]'

    def escape_outer_brackets(match):
        """Decide whether to escape the outer brackets."""
        inner = match.group(1)

        # Check if inner content has markdown links
        has_links = bool(re.search(r'\[[^\]]+\]\([^\)]+\)', inner))

        # Only escape if there are multiple links or commas (indicates a list)
        # inner.count('[') > 2 means at least 2 markdown links (each has 2 brackets)
        if has_links and (inner.count('[') > 2 or ',' in inner):
            return f'\\[{inner}\\]'

        return match.group(0)  # Don't modify single links

    # Apply once (recursive application can cause issues with already-escaped brackets)
    text = re.sub(pattern, escape_outer_brackets, text)
    return text


def html_to_markdown(html_text: str, highlight_syntax: str = '==') -> str:
    """Convert HTML to markdown with comprehensive tag and entity handling.

    This is a shared utility used by both Obsidian and Logseq converters.

    Args:
        html_text: HTML text to convert
        highlight_syntax: Syntax for highlighting ('==' for Obsidian, '^^' for Logseq)

    Returns:
        Markdown text with HTML converted

    Handles:
        - HTML entities (&nbsp;, &quot;, &lt;, &gt;, &amp;, etc.)
        - Bold: <strong>, <b>, <span style='font-weight:bold'>
        - Italic: <em>, <i>, <span style='font-style:italic'>
        - Underline: <u> (converted to bold for markdown compatibility)
        - Strikethrough: <s>, <strike>, <del>, <span style='text-decoration:line-through'>
        - Code: <code>, <tt>
        - Superscript: <sup> (markdown extension)
        - Subscript: <sub> (markdown extension)
        - Highlighting: <span style='background:yellow'>
        - Links: <a href="">
        - Line breaks: <br>
        - All other HTML tags are stripped but content is preserved
    """
    text = html_text

    # STEP 1: Decode HTML entities FIRST
    # This handles &nbsp;, &quot;, &lt;, &gt;, &amp;, &apos;, etc.
    text = html.unescape(text)

    # STEP 2: Convert formatting tags to markdown (order matters!)
    # Use DOTALL flag to handle multi-line tags

    # Bold - handle both tags and inline styles
    text = re.sub(r'<strong>(.*?)</strong>', r'**\1**', text, flags=re.DOTALL)
    text = re.sub(r'<b>(.*?)</b>', r'**\1**', text, flags=re.DOTALL)
    text = re.sub(r"<span\s+style='font-weight:bold'[^>]*>(.*?)</span>", r'**\1**', text, flags=re.DOTALL)
    text = re.sub(r'<span\s+style="font-weight:bold"[^>]*>(.*?)</span>', r'**\1**', text, flags=re.DOTALL)

    # Italic - handle both tags and inline styles
    text = re.sub(r'<em>(.*?)</em>', r'*\1*', text, flags=re.DOTALL)
    text = re.sub(r'<i>(.*?)</i>', r'*\1*', text, flags=re.DOTALL)
    text = re.sub(r"<span\s+style='font-style:italic'[^>]*>(.*?)</span>", r'*\1*', text, flags=re.DOTALL)
    text = re.sub(r'<span\s+style="font-style:italic"[^>]*>(.*?)</span>', r'*\1*', text, flags=re.DOTALL)

    # Underline - convert to bold (markdown doesn't have native underline)
    text = re.sub(r'<u>(.*?)</u>', r'**\1**', text, flags=re.DOTALL)
    text = re.sub(r"<span\s+style='text-decoration:underline'[^>]*>(.*?)</span>", r'**\1**', text, flags=re.DOTALL)
    text = re.sub(r'<span\s+style="text-decoration:underline"[^>]*>(.*?)</span>', r'**\1**', text, flags=re.DOTALL)

    # Strikethrough
    text = re.sub(r'<s>(.*?)</s>', r'~~\1~~', text, flags=re.DOTALL)
    text = re.sub(r'<strike>(.*?)</strike>', r'~~\1~~', text, flags=re.DOTALL)
    text = re.sub(r'<del>(.*?)</del>', r'~~\1~~', text, flags=re.DOTALL)
    text = re.sub(r"<span\s+style='text-decoration:line-through'[^>]*>(.*?)</span>", r'~~\1~~', text, flags=re.DOTALL)
    text = re.sub(r'<span\s+style="text-decoration:line-through"[^>]*>(.*?)</span>', r'~~\1~~', text, flags=re.DOTALL)

    # Code - inline code
    text = re.sub(r'<code>(.*?)</code>', r'`\1`', text, flags=re.DOTALL)
    text = re.sub(r'<tt>(.*?)</tt>', r'`\1`', text, flags=re.DOTALL)

    # Superscript and Subscript (markdown extensions)
    text = re.sub(r'<sup>(.*?)</sup>', r'^[\1]', text, flags=re.DOTALL)
    text = re.sub(r'<sub>(.*?)</sub>', r'~[\1]', text, flags=re.DOTALL)

    # Highlighting - OneNote yellow highlight
    # Use the provided highlight syntax (== for Obsidian, ^^ for Logseq)
    text = re.sub(
        r"<span\s+style='background:yellow;mso-highlight:yellow'[^>]*>(.*?)</span>",
        rf"{highlight_syntax}\1{highlight_syntax}",
        text,
        flags=re.DOTALL
    )
    text = re.sub(
        r'<span\s+style="background:yellow;mso-highlight:yellow"[^>]*>(.*?)</span>',
        rf"{highlight_syntax}\1{highlight_syntax}",
        text,
        flags=re.DOTALL
    )

    # Links
    text = re.sub(r'<a\s+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r'[\2](\1)', text, flags=re.DOTALL)

    # Line breaks
    text = re.sub(r'<br\s*/?>', '\n', text)

    # STEP 3: Strip remaining HTML tags (colors, fonts, language tags, etc.)
    # These don't have markdown equivalents, so we just keep the content
    text = re.sub(r'<[^>]+>', '', text)

    # STEP 4: Clean up whitespace
    text = re.sub(r'\n\s*\n', '\n\n', text)
    text = text.strip()

    return text
