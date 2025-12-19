"""Text utilities for wrapping and sizing in PD Generator."""

import logging
from typing import List, Optional, Tuple

from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth

logger = logging.getLogger(__name__)


def wrap_text(
    text: str,
    max_width: float,
    font_name: str,
    font_size: float,
) -> List[str]:
    """
    Wrap text to fit within a maximum width.

    Args:
        text: Text to wrap
        max_width: Maximum width in points
        font_name: Font name for measuring
        font_size: Font size in points

    Returns:
        List of wrapped lines
    """
    if not text:
        return []

    words = text.replace("\n", " \n ").split(" ")
    lines = []
    current_line = ""

    for word in words:
        if word == "\n":
            if current_line:
                lines.append(current_line.strip())
            current_line = ""
            continue

        if not word:
            continue

        test_line = f"{current_line} {word}".strip() if current_line else word
        test_width = stringWidth(test_line, font_name, font_size)

        if test_width <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            # Check if word itself is too long
            if stringWidth(word, font_name, font_size) > max_width:
                # Break the word
                lines.extend(_break_long_word(word, max_width, font_name, font_size))
                current_line = ""
            else:
                current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def _break_long_word(
    word: str,
    max_width: float,
    font_name: str,
    font_size: float,
) -> List[str]:
    """Break a long word into multiple lines."""
    result = []
    current = ""

    for char in word:
        test = current + char
        if stringWidth(test, font_name, font_size) <= max_width:
            current = test
        else:
            if current:
                result.append(current)
            current = char

    if current:
        result.append(current)

    return result


def calculate_text_height(
    lines: List[str],
    font_size: float,
    line_spacing: float = 1.2,
) -> float:
    """
    Calculate total height of wrapped text.

    Args:
        lines: List of text lines
        font_size: Font size in points
        line_spacing: Line spacing multiplier

    Returns:
        Total height in points
    """
    if not lines:
        return 0

    line_height = font_size * line_spacing
    return len(lines) * line_height


def fit_text_to_box(
    text: str,
    max_width: float,
    max_height: float,
    font_name: str,
    initial_font_size: float,
    min_font_size: float,
    line_spacing: float = 1.2,
) -> Tuple[List[str], float, bool]:
    """
    Fit text into a box, reducing font size if needed.

    Args:
        text: Text to fit
        max_width: Maximum width in points
        max_height: Maximum height in points
        font_name: Font name
        initial_font_size: Starting font size in points
        min_font_size: Minimum font size before truncation
        line_spacing: Line spacing multiplier

    Returns:
        Tuple of (lines, final_font_size, was_truncated)
    """
    if not text:
        return [], initial_font_size, False

    font_size = initial_font_size
    was_truncated = False

    while font_size >= min_font_size:
        lines = wrap_text(text, max_width, font_name, font_size)
        height = calculate_text_height(lines, font_size, line_spacing)

        if height <= max_height:
            return lines, font_size, False

        # Reduce font size and try again
        font_size -= 1

    # Font size at minimum but still doesn't fit - truncate
    font_size = min_font_size
    lines = wrap_text(text, max_width, font_name, font_size)
    line_height = font_size * line_spacing
    max_lines = int(max_height / line_height)

    if len(lines) > max_lines:
        lines = lines[:max_lines]
        # Add ellipsis to last line
        if lines:
            last_line = lines[-1]
            ellipsis = "…"
            # Shorten last line to fit ellipsis
            while last_line and stringWidth(last_line + ellipsis, font_name, font_size) > max_width:
                last_line = last_line[:-1]
            lines[-1] = last_line + ellipsis
        was_truncated = True

    return lines, font_size, was_truncated


def sanitize_filename(name: str, max_length: int = 200) -> str:
    """
    Create a safe filename from a string.

    Args:
        name: Original name
        max_length: Maximum filename length

    Returns:
        Safe filename string
    """
    # Replace problematic characters
    unsafe_chars = '<>:"/\\|?*\x00'
    result = name

    for char in unsafe_chars:
        result = result.replace(char, "_")

    # Replace multiple spaces/underscores with single underscore
    while "  " in result:
        result = result.replace("  ", " ")
    while "__" in result:
        result = result.replace("__", "_")

    # Strip leading/trailing spaces and dots
    result = result.strip(" .")

    # Truncate if too long (leave room for extension)
    if len(result) > max_length:
        result = result[:max_length]

    return result


def format_output_filename(
    pattern: str,
    project_id: str,
    project_name: str,
) -> str:
    """
    Format output filename from pattern.

    Args:
        pattern: Naming pattern with placeholders
        project_id: Project ID
        project_name: Project name

    Returns:
        Formatted filename (without extension)
    """
    result = pattern.replace("{project_id}", str(project_id))
    result = result.replace("{project_name}", project_name)
    return sanitize_filename(result)
