# PD Generator - University Project Poster Generator

A Windows-first CLI tool that automatically generates A1 posters (594×841 mm) for university projects from Excel data.

## Features

- **Batch PDF Generation**: Generate print-ready A1 posters for multiple projects
- **Cyrillic Support**: Full support for Russian and other Cyrillic text
- **Configurable Layout**: YAML/JSON configuration for fonts, sizes, positioning
- **Smart Text Handling**: Auto-wrap with font scaling and graceful truncation
- **Robust Error Handling**: Per-project errors don't stop the whole batch
- **Detailed Logging**: Clear progress and error reporting with summary

## Installation

### Prerequisites

- Python 3.8 or higher
- Windows (primary target), also works on Linux/macOS

### Install Dependencies

```bash
# Clone or download the repository
cd PD_generator

# Install using pip (recommended)
pip install -e .

# Or install dependencies directly
pip install -r requirements.txt
```

### Font Setup (for Cyrillic support)

The generator uses DejaVu fonts for Cyrillic text support. On Windows, these are usually pre-installed. If not:

1. Download DejaVu fonts from [dejavu-fonts.github.io](https://dejavu-fonts.github.io/)
2. Install the fonts by double-clicking the .ttf files

On Linux (Debian/Ubuntu):
```bash
sudo apt-get install fonts-dejavu
```

## Quick Start

1. **Prepare your Excel file** with columns:
   - `project_id` - Unique identifier
   - `project_name` - Project title
   - `problem` - Problem description
   - `solution` - Solution description
   - `product` - Product description
   - `team` - Team members (can use newlines)
   - `image_filename` - (Optional) Image filename

2. **Prepare project images** in a folder, named by project_id (e.g., `101.jpg`, `102.png`)

3. **Run the generator**:
   ```bash
   # Using installed command
   pd-generator projects.xlsx images/ --output posters/

   # Or using Python directly
   python main.py projects.xlsx images/ --output posters/
   ```

## Usage

### Basic Usage

```bash
pd-generator <excel_file> <images_folder> [options]
```

### Options

| Option | Description |
|--------|-------------|
| `-o, --output <folder>` | Output folder for PDFs (default: `output`) |
| `-c, --config <file>` | Configuration file (YAML/JSON) |
| `--only <id1> <id2> ...` | Generate only specific project IDs |
| `-v, --verbose` | Enable verbose logging |
| `-q, --quiet` | Suppress output except errors |
| `--init-config` | Print default configuration YAML |
| `--version` | Show version |
| `--help` | Show help message |

### Examples

```bash
# Generate all posters
pd-generator projects.xlsx images/

# Generate specific projects only
pd-generator projects.xlsx images/ --only 101 102 103

# Use custom configuration
pd-generator projects.xlsx images/ --config my_config.yaml

# Generate default config file
pd-generator --init-config > config.yaml

# Verbose output for debugging
pd-generator projects.xlsx images/ -v
```

## Configuration

Create a `config.yaml` file (or use `--init-config` to generate one):

```yaml
# Page dimensions (A1 portrait)
page:
  width_mm: 594
  height_mm: 841  # Use 840 if needed for your printer

# Layout configuration
layout:
  image_height_mm: 434           # Top image area height
  image_fit_mode: cover          # "cover" or "contain"
  content_padding_left_mm: 40
  content_padding_right_mm: 40
  content_padding_top_mm: 20
  content_padding_bottom_mm: 20
  text_column_width_mm: 225
  title_y_offset_mm: 50
  title_centered: true

# Font configuration
fonts:
  title_font: DejaVuSans-Bold
  title_size: 48
  heading_font: DejaVuSans-Bold
  heading_size: 24
  body_font: DejaVuSans
  body_size: 18
  min_font_size: 10              # Minimum before truncation
  line_spacing: 1.2

# Excel column mapping (customize if your columns have different names)
columns:
  project_id: project_id
  project_name: project_name
  problem: problem
  solution: solution
  product: product
  team: team
  image_filename: image_filename

# Output configuration
output:
  naming_pattern: "{project_id}_{project_name}"
  output_folder: output

# Logo configuration (optional)
logos:
  paths:
    - logos/university_logo.png
    - logos/faculty_logo.png
  height_mm: 40
  spacing_mm: 10
  position: bottom_left
```

## Poster Layout

```
┌─────────────────────────────────────────┐
│                                         │
│           TOP IMAGE AREA                │
│           (594 × 434 mm)                │
│                                         │
├─────────────────────────────────────────┤
│                                         │
│         [PROJECT TITLE - centered]      │
│                                         │
│  ┌──────────────┐  ┌──────────────────┐ │
│  │              │  │    PROBLEM       │ │
│  │    LOGOS     │  │    -----------   │ │
│  │   (optional) │  │    text...       │ │
│  │              │  │                  │ │
│  │              │  │    SOLUTION      │ │
│  │              │  │    -----------   │ │
│  │              │  │    text...       │ │
│  │              │  │                  │ │
│  │              │  │    PRODUCT       │ │
│  │              │  │    -----------   │ │
│  │              │  │    text...       │ │
│  │              │  │                  │ │
│  │              │  │    TEAM          │ │
│  │              │  │    -----------   │ │
│  └──────────────┘  └──────────────────┘ │
│     40mm padding      225mm text col    │
└─────────────────────────────────────────┘
```

## Excel Template

See `example_data/projects.xlsx` for a complete example. Create it by running:

```bash
cd example_data
python create_example_excel.py
```

### Required Columns

| Column | Description |
|--------|-------------|
| `project_id` | Unique identifier (used for image matching) |
| `project_name` | Project title displayed on poster |
| `problem` | Problem description text |
| `solution` | Solution description text |
| `product` | Product description text |
| `team` | Team members (use newlines for multiple) |

### Optional Columns

| Column | Description |
|--------|-------------|
| `image_filename` | Explicit image filename (if not using project_id naming) |

## Image Requirements

- **Formats**: JPG, JPEG, PNG, GIF, BMP, TIFF
- **Naming**: `{project_id}.jpg` or `{project_id}.png` (e.g., `101.jpg`)
- **Recommended size**: At least 1800×1300 pixels for good print quality
- **Aspect ratio**: Will be scaled to fit (cropped if using "cover" mode)

## Error Handling

The generator is designed to be robust:

- Missing images: Poster generated with placeholder
- Missing fields: Project skipped with error logged
- Text overflow: Font reduced, then truncated with "…" + warning
- Missing logos: Silently skipped
- Invalid rows: Skipped without stopping batch

## Output

- PDFs are saved to the output folder
- Filename pattern: `{project_id}_{project_name}.pdf`
- Special characters are sanitized for Windows compatibility
- Summary report shows success/failure counts

## Troubleshooting

### Fonts not rendering correctly

Ensure DejaVu fonts are installed. The generator falls back to Helvetica if DejaVu is not found (which doesn't support Cyrillic).

### Text appearing truncated

1. Check your text length
2. Increase `text_column_width_mm` in config
3. Reduce `body_size` or `min_font_size`

### Images not found

- Ensure image files are named exactly as the project_id (e.g., `101.jpg`)
- Or use the `image_filename` column in Excel
- Supported formats: .jpg, .jpeg, .png, .gif, .bmp, .tiff

## Project Structure

```
PD_generator/
├── src/
│   └── pd_generator/
│       ├── __init__.py      # Package initialization
│       ├── cli.py           # Command-line interface
│       ├── config.py        # Configuration handling
│       ├── excel_reader.py  # Excel file parsing
│       ├── poster.py        # PDF generation (ReportLab)
│       └── text_utils.py    # Text wrapping utilities
├── example_data/
│   └── create_example_excel.py
├── config.yaml              # Default configuration
├── main.py                  # Entry point
├── requirements.txt
├── setup.py
└── README.md
```

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.
