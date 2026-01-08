"""
Step 1: Convert all PowerPoint slides to images (Recursive)
- Extracts embedded OLE objects
- Converts all slides to images
- Saves metadata for Step 2
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Set
from datetime import datetime
from tqdm import tqdm

from config import DATA_DIR, OUTPUT_DIR, TEMP_DIR
from pptx_extractor import extract_embedded_pptx
from image_converter import convert_pptx_to_images

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class Step1Converter:
    """
    Step 1: Convert all PowerPoint files to images recursively
    """

    def __init__(
        self,
        input_dir: Path = DATA_DIR,
        output_dir: Path = OUTPUT_DIR,
        temp_dir: Path = TEMP_DIR
    ):
        """
        Initialize the converter

        Args:
            input_dir: Directory containing PowerPoint files
            output_dir: Directory for output files
            temp_dir: Directory for temporary files
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(temp_dir)
        self.processed_files: Set[str] = set()

        # Ensure directories exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        # Storage for file hierarchy and image mappings
        self.file_hierarchy: Dict = {}
        self.image_registry: Dict[str, List[str]] = {}  # Store as strings for JSON

    def process_all(self) -> Dict:
        """
        Main processing pipeline: convert all slides to images

        Returns:
            Dictionary containing metadata
        """
        if not self.input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {self.input_dir}")

        # Find all PowerPoint files
        pptx_files = list(self.input_dir.glob("*.pptx")) + list(self.input_dir.glob("*.ppt"))
        pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

        if not pptx_files:
            logger.warning(f"No PowerPoint files found in {self.input_dir}")
            return {}

        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 1: Converting all slides to images (recursive)")
        logger.info(f"Input directory: {self.input_dir}")
        logger.info(f"Found {len(pptx_files)} PowerPoint files")
        logger.info(f"{'='*80}\n")

        # Process all files recursively
        for pptx_file in pptx_files:
            logger.info(f"\nProcessing: {pptx_file.name}")
            self._recursive_convert(pptx_file, parent_file=None, depth=0)

        # Save metadata
        metadata = self._save_metadata()

        # Print summary
        total_images = sum(len(imgs) for imgs in self.image_registry.values())
        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 1 COMPLETED!")
        logger.info(f"Total files processed: {len(self.image_registry)}")
        logger.info(f"Total images converted: {total_images}")
        logger.info(f"Metadata saved to: {self.output_dir / 'step1_metadata.json'}")
        logger.info(f"Images saved to: {self.temp_dir}")
        logger.info(f"{'='*80}\n")

        return metadata

    def _recursive_convert(
        self,
        pptx_path: Path,
        parent_file: str = None,
        depth: int = 0
    ) -> None:
        """
        Recursively convert PowerPoint files to images

        Args:
            pptx_path: Path to PowerPoint file
            parent_file: Name of parent file
            depth: Current recursion depth
        """
        pptx_path = Path(pptx_path)
        file_key = f"{pptx_path.stem}_d{depth}"

        # Avoid processing the same file twice
        if file_key in self.processed_files:
            logger.info(f"{'  ' * depth}Skipping already processed: {pptx_path.name}")
            return

        self.processed_files.add(file_key)
        indent = '  ' * depth

        logger.info(f"{indent}[Depth {depth}] Processing: {pptx_path.name}")

        try:
            # Convert slides to images
            slide_images_dir = self.temp_dir / f"{file_key}_slides"
            logger.info(f"{indent}Converting slides to images...")

            slide_images = convert_pptx_to_images(
                pptx_path,
                slide_images_dir,
                show_progress=True
            )

            # Register images (store as strings for JSON serialization)
            self.image_registry[file_key] = [str(img) for img in slide_images]

            # Store file hierarchy
            self.file_hierarchy[file_key] = {
                "filename": pptx_path.name,
                "parent_file": parent_file,
                "depth": depth,
                "total_slides": len(slide_images),
                "file_path": str(pptx_path),
                "images_dir": str(slide_images_dir)
            }

            logger.info(f"{indent}Converted {len(slide_images)} slides")

            # Extract embedded OLE objects
            logger.info(f"{indent}Checking for embedded files...")
            ole_dir = self.temp_dir / f"{file_key}_embedded"
            embedded_pptx_files = extract_embedded_pptx(pptx_path, ole_dir)

            if embedded_pptx_files:
                logger.info(f"{indent}Found {len(embedded_pptx_files)} embedded PowerPoint files")
                self.file_hierarchy[file_key]["embedded_files"] = [
                    f.name for f in embedded_pptx_files
                ]

                # Recursively process embedded files
                for embedded_file in embedded_pptx_files:
                    self._recursive_convert(
                        embedded_file,
                        parent_file=pptx_path.name,
                        depth=depth + 1
                    )
            else:
                logger.info(f"{indent}No embedded files found")
                self.file_hierarchy[file_key]["embedded_files"] = []

        except Exception as e:
            logger.error(f"{indent}Error processing {pptx_path.name}: {e}", exc_info=True)
            self.file_hierarchy[file_key] = {
                "filename": pptx_path.name,
                "parent_file": parent_file,
                "depth": depth,
                "error": str(e)
            }

    def _save_metadata(self) -> Dict:
        """
        Save metadata to JSON file for Step 2

        Returns:
            Metadata dictionary
        """
        metadata = {
            "step1_info": {
                "total_files_processed": len(self.image_registry),
                "total_images_converted": sum(len(imgs) for imgs in self.image_registry.values()),
                "processed_at": datetime.now().isoformat(),
                "input_dir": str(self.input_dir),
                "temp_dir": str(self.temp_dir)
            },
            "file_hierarchy": self.file_hierarchy,
            "image_registry": self.image_registry
        }

        metadata_path = self.output_dir / "step1_metadata.json"
        with open(metadata_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)

        logger.info(f"Metadata saved to: {metadata_path}")
        return metadata


def main():
    """Main entry point for Step 1"""
    logger.info("="*80)
    logger.info("VLM Recursive OCR - STEP 1: Image Conversion")
    logger.info("="*80)

    converter = Step1Converter()

    try:
        metadata = converter.process_all()

        logger.info("\nNext step:")
        logger.info("  Run: python step2_analyze.py")

    except Exception as e:
        logger.error(f"Fatal error during Step 1: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
