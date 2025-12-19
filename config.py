"""Configuration handling for PD Generator."""

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml


@dataclass
class PageConfig:
    """Page dimension configuration."""
    width_mm: float = 594.0
    height_mm: float = 841.0  # Can be 840 or 841


@dataclass
class LayoutConfig:
    """Layout configuration for poster elements."""
    # Top image area
    image_height_mm: float = 434.0
    image_fit_mode: str = "cover"  # "cover" or "contain"

    # Bottom content area
    content_padding_left_mm: float = 40.0
    content_padding_right_mm: float = 40.0
    content_padding_top_mm: float = 20.0
    content_padding_bottom_mm: float = 20.0

    # Text column (bottom right)
    text_column_width_mm: float = 225.0

    # Title configuration
    title_y_offset_mm: float = 50.0  # Distance from top of content area
    title_centered: bool = True


@dataclass
class FontConfig:
    """Font configuration."""
    title_font: str = "DejaVuSans-Bold"
    title_size: float = 48.0

    heading_font: str = "DejaVuSans-Bold"
    heading_size: float = 24.0

    body_font: str = "DejaVuSans"
    body_size: float = 18.0

    min_font_size: float = 10.0
    line_spacing: float = 1.2


@dataclass
class ColumnMapping:
    """Mapping of Excel columns to poster fields."""
    project_id: str = "project_id"
    project_name: str = "project_name"
    problem: str = "problem"
    solution: str = "solution"
    product: str = "product"
    team: str = "team"
    image_filename: Optional[str] = "image_filename"


@dataclass
class OutputConfig:
    """Output configuration."""
    naming_pattern: str = "{project_id}_{project_name}"
    output_folder: str = "output"


@dataclass
class LogoConfig:
    """Logo configuration."""
    paths: List[str] = field(default_factory=list)
    height_mm: float = 40.0
    spacing_mm: float = 10.0
    position: str = "bottom_left"  # Position in the content area
    margin_left_mm: float = 10.0   # Left margin from page edge
    margin_bottom_mm: float = 10.0  # Bottom margin from page edge


@dataclass
class Config:
    """Main configuration container."""
    page: PageConfig = field(default_factory=PageConfig)
    layout: LayoutConfig = field(default_factory=LayoutConfig)
    fonts: FontConfig = field(default_factory=FontConfig)
    columns: ColumnMapping = field(default_factory=ColumnMapping)
    output: OutputConfig = field(default_factory=OutputConfig)
    logos: LogoConfig = field(default_factory=LogoConfig)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Config":
        """Create Config from dictionary."""
        config = cls()

        if "page" in data:
            page_data = data["page"]
            config.page = PageConfig(
                width_mm=page_data.get("width_mm", 594.0),
                height_mm=page_data.get("height_mm", 841.0),
            )

        if "layout" in data:
            layout_data = data["layout"]
            config.layout = LayoutConfig(
                image_height_mm=layout_data.get("image_height_mm", 434.0),
                image_fit_mode=layout_data.get("image_fit_mode", "cover"),
                content_padding_left_mm=layout_data.get("content_padding_left_mm", 40.0),
                content_padding_right_mm=layout_data.get("content_padding_right_mm", 40.0),
                content_padding_top_mm=layout_data.get("content_padding_top_mm", 20.0),
                content_padding_bottom_mm=layout_data.get("content_padding_bottom_mm", 20.0),
                text_column_width_mm=layout_data.get("text_column_width_mm", 225.0),
                title_y_offset_mm=layout_data.get("title_y_offset_mm", 50.0),
                title_centered=layout_data.get("title_centered", True),
            )

        if "fonts" in data:
            fonts_data = data["fonts"]
            config.fonts = FontConfig(
                title_font=fonts_data.get("title_font", "DejaVuSans-Bold"),
                title_size=fonts_data.get("title_size", 48.0),
                heading_font=fonts_data.get("heading_font", "DejaVuSans-Bold"),
                heading_size=fonts_data.get("heading_size", 24.0),
                body_font=fonts_data.get("body_font", "DejaVuSans"),
                body_size=fonts_data.get("body_size", 18.0),
                min_font_size=fonts_data.get("min_font_size", 10.0),
                line_spacing=fonts_data.get("line_spacing", 1.2),
            )

        if "columns" in data:
            cols_data = data["columns"]
            config.columns = ColumnMapping(
                project_id=cols_data.get("project_id", "project_id"),
                project_name=cols_data.get("project_name", "project_name"),
                problem=cols_data.get("problem", "problem"),
                solution=cols_data.get("solution", "solution"),
                product=cols_data.get("product", "product"),
                team=cols_data.get("team", "team"),
                image_filename=cols_data.get("image_filename", "image_filename"),
            )

        if "output" in data:
            output_data = data["output"]
            config.output = OutputConfig(
                naming_pattern=output_data.get("naming_pattern", "{project_id}_{project_name}"),
                output_folder=output_data.get("output_folder", "output"),
            )

        if "logos" in data:
            logos_data = data["logos"]
            config.logos = LogoConfig(
                paths=logos_data.get("paths", []),
                height_mm=logos_data.get("height_mm", 40.0),
                spacing_mm=logos_data.get("spacing_mm", 10.0),
                position=logos_data.get("position", "bottom_left"),
                margin_left_mm=logos_data.get("margin_left_mm", 10.0),
                margin_bottom_mm=logos_data.get("margin_bottom_mm", 10.0),
            )

        return config

    @classmethod
    def from_yaml(cls, path: Path) -> "Config":
        """Load configuration from YAML file."""
        if not path.exists():
            return cls()

        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

        return cls.from_dict(data)

    @classmethod
    def from_json(cls, path: Path) -> "Config":
        """Load configuration from JSON file."""
        import json

        if not path.exists():
            return cls()

        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        return cls.from_dict(data)

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "Config":
        """Load configuration from file (auto-detect format)."""
        if path is None:
            # Try default locations
            for default_path in [Path("config.yaml"), Path("config.yml"), Path("config.json")]:
                if default_path.exists():
                    path = default_path
                    break

        if path is None or not path.exists():
            return cls()

        suffix = path.suffix.lower()
        if suffix in (".yaml", ".yml"):
            return cls.from_yaml(path)
        elif suffix == ".json":
            return cls.from_json(path)
        else:
            raise ValueError(f"Unsupported config file format: {suffix}")


def get_default_config_yaml() -> str:
    """Return default configuration as YAML string."""
    return """# PD Generator Configuration

# Page dimensions (A1 portrait)
page:
  width_mm: 594
  height_mm: 841  # Use 840 if needed for your printer

# Layout configuration
layout:
  # Top image area
  image_height_mm: 434
  image_fit_mode: cover  # "cover" (fill, may crop) or "contain" (fit, may have margins)

  # Bottom content area padding
  content_padding_left_mm: 40
  content_padding_right_mm: 40
  content_padding_top_mm: 20
  content_padding_bottom_mm: 20

  # Text column (bottom right area)
  text_column_width_mm: 225

  # Title positioning
  title_y_offset_mm: 50  # Distance from top of content area
  title_centered: true

# Font configuration
fonts:
  title_font: DejaVuSans-Bold
  title_size: 48
  heading_font: DejaVuSans-Bold
  heading_size: 24
  body_font: DejaVuSans
  body_size: 18
  min_font_size: 10  # Minimum font size before truncation
  line_spacing: 1.2

# Excel column mapping
columns:
  project_id: project_id
  project_name: project_name
  problem: problem
  solution: solution
  product: product
  team: team
  image_filename: image_filename  # Optional: if not present, uses project_id to find image

# Output configuration
output:
  naming_pattern: "{project_id}_{project_name}"
  output_folder: output

# Logo configuration (optional)
logos:
  paths:
    # - logos/university_logo.png
    # - logos/faculty_logo.png
  height_mm: 40
  spacing_mm: 10
  position: bottom_left
"""
