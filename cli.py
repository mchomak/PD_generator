"""Auto-run entry point for PD Generator (no CLI args)."""

import logging
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from config import Config
from excel_reader import ExcelReader
from poster import generate_all_posters

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")
PREFERRED_IMAGES_DIR_NAMES = ("images", "img", "pictures", "pics")

# The app will ONLY use this exact Excel filename from the same directory as the executable
EXPECTED_EXCEL_NAME = "project_info.xlsx"


def _base_dir() -> Path:
    """
    Directory where the script/executable is located.
    Works for normal python run and for frozen executables (PyInstaller, etc.).
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "executable"):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _find_excel_file(folder: Path) -> Optional[Path]:
    """
    Find the Excel file ONLY by the expected fixed name in the same directory
    as the script/executable: project_info.xlsx
    """
    p = folder / EXPECTED_EXCEL_NAME
    if not p.exists() or not p.is_file():
        return None

    # Ignore temporary/metadata files (just in case)
    if p.name.startswith("~$") or p.name.startswith("._") or p.name.startswith("."):
        return None

    # Quick validation: .xlsx is a ZIP container (signature starts with PK\x03\x04)
    try:
        with p.open("rb") as f:
            sig = f.read(4)
        if sig != b"PK\x03\x04":
            logger.error("Excel file '%s' exists but is not a valid .xlsx (missing ZIP signature).", p.name)
            return None
    except Exception as e:
        logger.warning("Could not validate Excel file '%s': %s", p.name, e)

    return p


def _find_images_folder(folder: Path) -> Optional[Path]:
    # 1) Preferred conventional names
    for name in PREFERRED_IMAGES_DIR_NAMES:
        d = folder / name
        if d.exists() and d.is_dir():
            # Validate it has at least one image (optional, but useful)
            if any(p.suffix.lower() in IMAGE_EXTS for p in d.iterdir() if p.is_file()):
                return d
            # If folder exists but empty, still accept it (user may add images later)
            return d

    # 2) Heuristic: any subfolder that contains images
    candidates: List[Tuple[Path, int]] = []
    for d in folder.iterdir():
        if not d.is_dir():
            continue
        if d.name.lower() in ("output", "__pycache__", ".venv", "venv"):
            continue
        count = 0
        for p in d.iterdir():
            if p.is_file() and p.suffix.lower() in IMAGE_EXTS:
                count += 1
        if count > 0:
            candidates.append((d, count))

    if not candidates:
        return None

    # Pick the folder with the most images
    candidates.sort(key=lambda x: x[1], reverse=True)
    picked = candidates[0][0]

    if len(candidates) > 1:
        logger.warning(
            "Multiple image folders detected. Using: %s. Others: %s",
            picked.name,
            ", ".join(d.name for d, _ in candidates[1:]),
        )
    return picked


def _find_config_path(folder: Path) -> Optional[Path]:
    for name in ("config.yaml", "config.yml", "config.json"):
        p = folder / name
        if p.exists() and p.is_file():
            return p
    return None


def print_summary(successes: List[tuple], failures: List[tuple], output_folder: Path) -> None:
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
        print("\nAll posters generated successfully.")
    elif successes:
        print(f"\n{len(failures)} poster(s) failed to generate.")
    else:
        print("\nNo posters were generated.")

    print("=" * 60 + "\n")


def main() -> int:
    """Main entry point (no arguments)."""
    base = _base_dir()

    excel_file = _find_excel_file(base)
    if not excel_file:
        logger.error("Excel file '%s' not found in: %s", EXPECTED_EXCEL_NAME, base)
        logger.error("Put '%s' in the same folder as the .exe and rerun.", EXPECTED_EXCEL_NAME)
        return 1

    images_folder = _find_images_folder(base)
    if not images_folder:
        logger.error("Images folder not found in: %s", base)
        logger.error(
            "Create a folder named 'images' next to the .exe (or any subfolder with .jpg/.png files) and rerun."
        )
        return 1

    output_folder = base / "output"
    output_folder.mkdir(parents=True, exist_ok=True)

    config_path = _find_config_path(base)

    # Load configuration
    try:
        config = Config.load(config_path)
        logger.info("Configuration loaded%s", f" from {config_path.name}" if config_path else " (default)")
    except Exception as e:
        logger.error("Failed to load configuration: %s", e)
        return 1

    # Force output folder to local ./output
    config.output.output_folder = str(output_folder)

    # Read Excel file
    logger.info("Reading Excel file: %s", excel_file.name)
    reader = ExcelReader(excel_file, config.columns)
    projects, global_errors = reader.read_projects()

    if global_errors:
        for error in global_errors:
            logger.error(error)
        return 1

    if not projects:
        logger.error("No projects found in Excel file: %s", excel_file.name)
        return 1

    logger.info("Found %d projects", len(projects))
    logger.info("Using images folder: %s", images_folder)
    logger.info("Generating posters to: %s", output_folder)

    successes, failures = generate_all_posters(
        projects=projects,
        config=config,
        images_folder=images_folder,
        output_folder=output_folder,
        only_ids=None,
    )

    print_summary(successes, failures, output_folder)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
