"""Command-line interface for PD Generator."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from . import __version__
from .config import Config, get_default_config_yaml
from .excel_reader import ExcelReader
from .poster import generate_all_posters

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def parse_args(args: Optional[List[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        prog="pd-generator",
        description="Generate A1 posters for university projects from Excel data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  pd-generator projects.xlsx images/
  pd-generator projects.xlsx images/ --output posters/
  pd-generator projects.xlsx images/ --only 101 102 103
  pd-generator projects.xlsx images/ --config custom_config.yaml
  pd-generator --init-config > config.yaml
        """,
    )

    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    parser.add_argument(
        "--init-config",
        action="store_true",
        help="Print default configuration YAML and exit",
    )

    parser.add_argument(
        "excel_file",
        nargs="?",
        type=Path,
        help="Path to Excel file with project data",
    )

    parser.add_argument(
        "images_folder",
        nargs="?",
        type=Path,
        help="Path to folder containing project images",
    )

    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=Path("output"),
        help="Output folder for generated PDFs (default: output)",
    )

    parser.add_argument(
        "-c", "--config",
        type=Path,
        help="Path to configuration file (YAML or JSON)",
    )

    parser.add_argument(
        "--only",
        nargs="+",
        metavar="PROJECT_ID",
        help="Generate posters only for specified project IDs",
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress all output except errors",
    )

    return parser.parse_args(args)


def print_summary(
    successes: List[tuple],
    failures: List[tuple],
    output_folder: Path,
):
    """Print a summary report of the generation process."""
    print("\n" + "=" * 60)
    print("GENERATION SUMMARY")
    print("=" * 60)

    total = len(successes) + len(failures)
    print(f"\nTotal projects processed: {total}")
    print(f"Successful: {len(successes)}")
    print(f"Failed: {len(failures)}")

    if successes:
        print(f"\nOutput folder: {output_folder.absolute()}")

    if failures:
        print("\n" + "-" * 40)
        print("FAILURES:")
        print("-" * 40)
        for project_id, error in failures:
            print(f"  [{project_id}] {error}")

    if successes and not failures:
        print("\n✓ All posters generated successfully!")
    elif successes:
        print(f"\n⚠ {len(failures)} poster(s) failed to generate")
    else:
        print("\n✗ No posters were generated")

    print("=" * 60 + "\n")


def main(args: Optional[List[str]] = None) -> int:
    """Main entry point for the CLI."""
    parsed_args = parse_args(args)

    # Handle --init-config
    if parsed_args.init_config:
        print(get_default_config_yaml())
        return 0

    # Configure logging level
    if parsed_args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    elif parsed_args.quiet:
        logging.getLogger().setLevel(logging.ERROR)

    # Validate required arguments
    if not parsed_args.excel_file:
        logger.error("Excel file is required. Use --help for usage information.")
        return 1

    if not parsed_args.images_folder:
        logger.error("Images folder is required. Use --help for usage information.")
        return 1

    excel_file = parsed_args.excel_file
    images_folder = parsed_args.images_folder
    output_folder = parsed_args.output

    # Validate inputs
    if not excel_file.exists():
        logger.error(f"Excel file not found: {excel_file}")
        return 1

    if not images_folder.exists():
        logger.error(f"Images folder not found: {images_folder}")
        return 1

    if not images_folder.is_dir():
        logger.error(f"Images path is not a folder: {images_folder}")
        return 1

    # Load configuration
    try:
        config = Config.load(parsed_args.config)
        logger.info("Configuration loaded")
    except Exception as e:
        logger.error(f"Failed to load configuration: {e}")
        return 1

    # Override output folder if specified in args
    if parsed_args.output:
        config.output.output_folder = str(parsed_args.output)
        output_folder = parsed_args.output

    # Read Excel file
    logger.info(f"Reading Excel file: {excel_file}")
    reader = ExcelReader(excel_file, config.columns)
    projects, global_errors = reader.read_projects()

    if global_errors:
        for error in global_errors:
            logger.error(error)
        return 1

    if not projects:
        logger.error("No projects found in Excel file")
        return 1

    logger.info(f"Found {len(projects)} projects")

    # Filter by project IDs if specified
    only_ids = parsed_args.only
    if only_ids:
        logger.info(f"Filtering to project IDs: {', '.join(only_ids)}")
        # Convert to strings for comparison
        only_ids = [str(pid) for pid in only_ids]

    # Generate posters
    logger.info(f"Generating posters to: {output_folder}")
    successes, failures = generate_all_posters(
        projects,
        config,
        images_folder,
        output_folder,
        only_ids,
    )

    # Print summary
    if not parsed_args.quiet:
        print_summary(successes, failures, output_folder)

    # Return non-zero if any failures
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
