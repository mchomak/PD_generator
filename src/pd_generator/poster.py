"""Poster generation with ReportLab for PD Generator."""

import logging
import os
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image
from reportlab.lib.colors import black, white
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from .config import Config
from .excel_reader import ProjectData
from .text_utils import fit_text_to_box, format_output_filename

logger = logging.getLogger(__name__)

# Standard font directories for different platforms
FONT_SEARCH_PATHS = [
    # Windows
    Path("C:/Windows/Fonts"),
    # Linux
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
    Path(os.path.expanduser("~/.fonts")),
    Path(os.path.expanduser("~/.local/share/fonts")),
    # macOS
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts"),
    Path(os.path.expanduser("~/Library/Fonts")),
]

# DejaVu font files (support Cyrillic)
DEJAVU_FONTS = {
    "DejaVuSans": "DejaVuSans.ttf",
    "DejaVuSans-Bold": "DejaVuSans-Bold.ttf",
    "DejaVuSans-Oblique": "DejaVuSans-Oblique.ttf",
    "DejaVuSans-BoldOblique": "DejaVuSans-BoldOblique.ttf",
    "DejaVuSerif": "DejaVuSerif.ttf",
    "DejaVuSerif-Bold": "DejaVuSerif-Bold.ttf",
}


def _find_font_file(font_filename: str) -> Optional[Path]:
    """Search for a font file in standard locations."""
    for search_path in FONT_SEARCH_PATHS:
        if not search_path.exists():
            continue

        # Search recursively
        for font_path in search_path.rglob(font_filename):
            if font_path.is_file():
                return font_path

    return None


def _register_fonts() -> bool:
    """Register fonts with ReportLab, return True if successful."""
    registered_any = False

    for font_name, font_filename in DEJAVU_FONTS.items():
        try:
            # Check if already registered
            pdfmetrics.getFont(font_name)
            registered_any = True
            continue
        except KeyError:
            pass

        font_path = _find_font_file(font_filename)
        if font_path:
            try:
                pdfmetrics.registerFont(TTFont(font_name, str(font_path)))
                logger.debug(f"Registered font: {font_name} from {font_path}")
                registered_any = True
            except Exception as e:
                logger.warning(f"Failed to register font {font_name}: {e}")

    return registered_any


def _ensure_fonts_available(config: Config) -> Tuple[str, str, str]:
    """
    Ensure required fonts are available, return actual font names to use.

    Returns:
        Tuple of (title_font, heading_font, body_font)
    """
    _register_fonts()

    # Try configured fonts first, fall back to Helvetica
    fonts = []
    for font_name in [config.fonts.title_font, config.fonts.heading_font, config.fonts.body_font]:
        try:
            pdfmetrics.getFont(font_name)
            fonts.append(font_name)
        except KeyError:
            logger.warning(f"Font '{font_name}' not available, using Helvetica")
            fonts.append("Helvetica")

    return tuple(fonts)


class PosterGenerator:
    """Generator for project posters."""

    def __init__(self, config: Config, images_folder: Path, output_folder: Path):
        """
        Initialize poster generator.

        Args:
            config: Configuration object
            images_folder: Folder containing project images
            output_folder: Folder for output PDFs
        """
        self.config = config
        self.images_folder = images_folder
        self.output_folder = output_folder
        self.warnings: List[str] = []

        # Ensure output folder exists
        self.output_folder.mkdir(parents=True, exist_ok=True)

        # Get actual fonts to use
        self.title_font, self.heading_font, self.body_font = _ensure_fonts_available(config)

    def _add_warning(self, message: str):
        """Add a warning message."""
        self.warnings.append(message)
        logger.warning(message)

    def _draw_image(
        self,
        c: canvas.Canvas,
        image_path: Path,
        x: float,
        y: float,
        width: float,
        height: float,
        fit_mode: str = "cover",
    ):
        """
        Draw an image on the canvas with specified fit mode.

        Args:
            c: ReportLab canvas
            image_path: Path to image file
            x: X position (bottom-left) in points
            y: Y position (bottom-left) in points
            width: Target width in points
            height: Target height in points
            fit_mode: "cover" (fill, may crop) or "contain" (fit within)
        """
        try:
            img = Image.open(image_path)
            img_width, img_height = img.size
            img_aspect = img_width / img_height
            target_aspect = width / height

            if fit_mode == "cover":
                # Scale to fill the entire area, cropping if needed
                if img_aspect > target_aspect:
                    # Image is wider - scale by height
                    scale = height / img_height
                    scaled_width = img_width * scale
                    scaled_height = height
                    # Center horizontally (will be cropped)
                    draw_x = x - (scaled_width - width) / 2
                    draw_y = y
                else:
                    # Image is taller - scale by width
                    scale = width / img_width
                    scaled_width = width
                    scaled_height = img_height * scale
                    # Center vertically (will be cropped)
                    draw_x = x
                    draw_y = y - (scaled_height - height) / 2

                # Save state, clip, draw, restore
                c.saveState()
                path = c.beginPath()
                path.rect(x, y, width, height)
                path.close()
                c.clipPath(path, stroke=0, fill=0)
                c.drawImage(str(image_path), draw_x, draw_y, scaled_width, scaled_height,
                           preserveAspectRatio=False)
                c.restoreState()
            else:
                # Contain mode - fit within bounds
                if img_aspect > target_aspect:
                    # Image is wider - scale by width
                    scaled_width = width
                    scaled_height = width / img_aspect
                else:
                    # Image is taller - scale by height
                    scaled_height = height
                    scaled_width = height * img_aspect

                # Center the image
                draw_x = x + (width - scaled_width) / 2
                draw_y = y + (height - scaled_height) / 2

                c.drawImage(str(image_path), draw_x, draw_y, scaled_width, scaled_height,
                           preserveAspectRatio=True)

        except Exception as e:
            self._add_warning(f"Failed to draw image {image_path}: {e}")
            # Draw placeholder rectangle
            c.setStrokeColor(black)
            c.setFillColor(white)
            c.rect(x, y, width, height, fill=1, stroke=1)

    def _draw_logos(self, c: canvas.Canvas, x: float, y: float, max_width: float, max_height: float):
        """Draw logos in the specified area."""
        if not self.config.logos.paths:
            return

        logo_height = self.config.logos.height_mm * mm
        spacing = self.config.logos.spacing_mm * mm
        current_x = x

        for logo_path_str in self.config.logos.paths:
            logo_path = Path(logo_path_str)
            if not logo_path.exists():
                logger.debug(f"Logo not found: {logo_path}")
                continue

            try:
                img = Image.open(logo_path)
                img_width, img_height = img.size
                aspect = img_width / img_height
                logo_width = logo_height * aspect

                if current_x + logo_width > x + max_width:
                    break  # No more space

                c.drawImage(
                    str(logo_path),
                    current_x,
                    y,
                    logo_width,
                    logo_height,
                    preserveAspectRatio=True,
                )
                current_x += logo_width + spacing

            except Exception as e:
                logger.debug(f"Failed to draw logo {logo_path}: {e}")

    def _draw_text_block(
        self,
        c: canvas.Canvas,
        text: str,
        x: float,
        y: float,
        max_width: float,
        max_height: float,
        font_name: str,
        font_size: float,
        heading: Optional[str] = None,
    ) -> Tuple[float, bool]:
        """
        Draw a text block with optional heading.

        Args:
            c: ReportLab canvas
            text: Text content
            x: X position (left) in points
            y: Y position (top) in points
            max_width: Maximum width in points
            max_height: Maximum height in points
            font_name: Font name for body text
            font_size: Font size for body text
            heading: Optional heading text

        Returns:
            Tuple of (height used, was_truncated)
        """
        line_spacing = self.config.fonts.line_spacing
        min_font_size = self.config.fonts.min_font_size
        heading_font = self.heading_font
        heading_size = self.config.fonts.heading_size

        current_y = y
        total_height = 0
        was_truncated = False

        # Draw heading if provided
        if heading:
            c.setFont(heading_font, heading_size)
            c.setFillColor(black)
            c.drawString(x, current_y - heading_size, heading)
            heading_height = heading_size * line_spacing * 1.5
            current_y -= heading_height
            total_height += heading_height
            max_height -= heading_height

        # Fit and draw body text
        lines, final_size, truncated = fit_text_to_box(
            text,
            max_width,
            max_height,
            font_name,
            font_size,
            min_font_size,
            line_spacing,
        )

        was_truncated = truncated

        c.setFont(font_name, final_size)
        line_height = final_size * line_spacing

        for line in lines:
            c.drawString(x, current_y - final_size, line)
            current_y -= line_height
            total_height += line_height

        return total_height, was_truncated

    def generate_poster(self, project: ProjectData, image_path: Optional[Path]) -> Tuple[Path, List[str]]:
        """
        Generate a poster for a project.

        Args:
            project: Project data
            image_path: Path to project image (optional)

        Returns:
            Tuple of (output path, list of warnings)
        """
        self.warnings = []

        # Calculate page dimensions
        page_width = self.config.page.width_mm * mm
        page_height = self.config.page.height_mm * mm

        # Calculate layout
        image_height = self.config.layout.image_height_mm * mm
        content_height = page_height - image_height

        padding_left = self.config.layout.content_padding_left_mm * mm
        padding_right = self.config.layout.content_padding_right_mm * mm
        padding_top = self.config.layout.content_padding_top_mm * mm
        padding_bottom = self.config.layout.content_padding_bottom_mm * mm

        text_column_width = self.config.layout.text_column_width_mm * mm
        logo_area_width = page_width - padding_left - padding_right - text_column_width - 20 * mm

        # Generate output filename
        filename = format_output_filename(
            self.config.output.naming_pattern,
            project.project_id,
            project.project_name,
        )
        output_path = self.output_folder / f"{filename}.pdf"

        # Create canvas
        c = canvas.Canvas(str(output_path), pagesize=(page_width, page_height))

        # Draw top image area
        image_y = page_height - image_height
        if image_path and image_path.exists():
            self._draw_image(
                c,
                image_path,
                0,
                image_y,
                page_width,
                image_height,
                self.config.layout.image_fit_mode,
            )
        else:
            # Draw placeholder
            c.setFillColor(white)
            c.setStrokeColor(black)
            c.rect(0, image_y, page_width, image_height, fill=1, stroke=1)
            c.setFillColor(black)
            c.setFont(self.body_font, 24)
            no_image_text = "No image available" if not image_path else f"Image not found: {image_path.name}"
            c.drawCentredString(page_width / 2, image_y + image_height / 2, no_image_text)

        # Draw project title (centered in content area)
        title_y_offset = self.config.layout.title_y_offset_mm * mm
        title_y = image_y - title_y_offset
        c.setFont(self.title_font, self.config.fonts.title_size)
        c.setFillColor(black)

        if self.config.layout.title_centered:
            c.drawCentredString(page_width / 2, title_y, project.project_name)
        else:
            c.drawString(padding_left, title_y, project.project_name)

        # Calculate text areas
        text_start_y = title_y - self.config.fonts.title_size * 1.5
        text_area_height = text_start_y - padding_bottom
        text_x = page_width - padding_right - text_column_width

        # Divide text area for three sections (problem, solution, product)
        section_height = (text_area_height - 60 * mm) / 3  # Leave some spacing

        # Draw text sections
        current_y = text_start_y
        sections = [
            ("PROBLEM", project.problem),
            ("SOLUTION", project.solution),
            ("PRODUCT", project.product),
        ]

        for heading, content in sections:
            height_used, truncated = self._draw_text_block(
                c,
                content,
                text_x,
                current_y,
                text_column_width,
                section_height,
                self.body_font,
                self.config.fonts.body_size,
                heading=heading,
            )
            if truncated:
                self._add_warning(f"Text truncated in {heading} section for project {project.project_id}")
            current_y -= height_used + 15 * mm

        # Draw team info at the bottom of text column
        team_y = padding_bottom + 30 * mm
        c.setFont(self.heading_font, self.config.fonts.heading_size)
        c.drawString(text_x, team_y + self.config.fonts.heading_size, "TEAM")
        c.setFont(self.body_font, self.config.fonts.body_size - 2)

        team_lines = project.team.split("\n") if project.team else []
        team_line_y = team_y
        for line in team_lines[:5]:  # Limit to 5 lines
            c.drawString(text_x, team_line_y, line.strip())
            team_line_y -= self.config.fonts.body_size * 1.2

        # Draw logos in bottom-left area
        logo_x = padding_left
        logo_y = padding_bottom
        self._draw_logos(c, logo_x, logo_y, logo_area_width, self.config.logos.height_mm * mm)

        # Save PDF
        c.save()

        return output_path, self.warnings


def generate_all_posters(
    projects: List[ProjectData],
    config: Config,
    images_folder: Path,
    output_folder: Path,
    only_ids: Optional[List[str]] = None,
) -> Tuple[List[Tuple[str, Path]], List[Tuple[str, str]]]:
    """
    Generate posters for all projects.

    Args:
        projects: List of project data
        config: Configuration object
        images_folder: Folder containing project images
        output_folder: Folder for output PDFs
        only_ids: Optional list of project IDs to generate (if None, generate all)

    Returns:
        Tuple of (success list, failure list)
        Success list: [(project_id, output_path), ...]
        Failure list: [(project_id, error_message), ...]
    """
    from .excel_reader import find_project_image

    generator = PosterGenerator(config, images_folder, output_folder)

    successes = []
    failures = []

    for project in projects:
        # Filter by ID if specified
        if only_ids is not None and project.project_id not in only_ids:
            continue

        try:
            # Validate project data
            errors = project.validate()
            if errors:
                failures.append((project.project_id, f"Validation errors: {'; '.join(errors)}"))
                continue

            # Find image
            image_path = find_project_image(project, images_folder)
            if image_path is None:
                logger.warning(f"No image found for project {project.project_id}")

            # Generate poster
            output_path, warnings = generator.generate_poster(project, image_path)

            if warnings:
                for warning in warnings:
                    logger.warning(f"Project {project.project_id}: {warning}")

            successes.append((project.project_id, output_path))
            logger.info(f"Generated poster: {output_path}")

        except Exception as e:
            error_msg = f"Failed to generate poster: {e}"
            failures.append((project.project_id, error_msg))
            logger.error(f"Project {project.project_id}: {error_msg}")

    return successes, failures
