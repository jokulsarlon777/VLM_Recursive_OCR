"""
VLM Recursive OCR - Main Script (2-Step Process)
Step 1: Recursively extract embedded files and convert all slides to images
Step 2: Analyze all images in parallel using VLM
"""
import json
import logging
import shutil
from pathlib import Path
from typing import Dict, List, Set, Tuple
from datetime import datetime
from tqdm import tqdm

from config import DATA_DIR, OUTPUT_DIR, TEMP_DIR
from pptx_extractor import extract_embedded_pptx
from image_converter import convert_pptx_to_images
from vlm_analyzer import analyze_slides

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class TwoStepPPTXProcessor:
    """
    Two-step PowerPoint processor:
    1. Convert all slides to images (including recursive extraction)
    2. Analyze all images in parallel using VLM
    """

    def __init__(
        self,
        output_dir: Path = OUTPUT_DIR,
        temp_dir: Path = TEMP_DIR,
        max_vlm_workers: int = 5
    ):
        """
        Initialize the processor

        Args:
            output_dir: Directory for output JSON files
            temp_dir: Directory for temporary files
            max_vlm_workers: Maximum number of parallel workers for VLM analysis
        """
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(temp_dir)
        self.max_vlm_workers = max_vlm_workers
        self.processed_files: Set[str] = set()

        # Ensure directories exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        # Storage for file hierarchy and image mappings
        self.file_hierarchy: Dict = {}
        self.image_registry: Dict[str, List[Path]] = {}

    def process_directory(self, input_dir: Path = DATA_DIR) -> Dict:
        """
        Main processing pipeline: two-step process

        Args:
            input_dir: Directory containing PowerPoint files

        Returns:
            Dictionary containing all results
        """
        input_dir = Path(input_dir)

        if not input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {input_dir}")

        # Find all PowerPoint files
        pptx_files = list(input_dir.glob("*.pptx")) + list(input_dir.glob("*.ppt"))
        pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

        if not pptx_files:
            logger.warning(f"No PowerPoint files found in {input_dir}")
            return {}

        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 1: Converting all slides to images (recursive)")
        logger.info(f"Found {len(pptx_files)} PowerPoint files")
        logger.info(f"{'='*80}\n")

        # Step 1: Convert all slides to images recursively
        for pptx_file in pptx_files:
            logger.info(f"\nProcessing: {pptx_file.name}")
            self._recursive_convert(pptx_file, parent_file=None, depth=0)

        # Step 2: Analyze all images in parallel
        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 2: Analyzing all images with VLM (parallel processing)")
        logger.info(f"{'='*80}\n")

        all_results = self._analyze_all_images()

        # Step 3: Generate output JSON files
        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 3: Generating JSON output files")
        logger.info(f"{'='*80}\n")

        self._generate_output_files(all_results)

        # Generate summary
        summary = {
            "processing_summary": {
                "total_files_processed": len(self.image_registry),
                "total_images_analyzed": sum(len(imgs) for imgs in self.image_registry.values()),
                "processed_at": datetime.now().isoformat()
            },
            "file_hierarchy": self.file_hierarchy,
            "results": all_results
        }

        summary_path = self.output_dir / "processing_summary.json"
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

        logger.info(f"\nSaved processing summary to: {summary_path}")
        return summary

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

            # Register images
            self.image_registry[file_key] = slide_images

            # Store file hierarchy
            self.file_hierarchy[file_key] = {
                "filename": pptx_path.name,
                "parent_file": parent_file,
                "depth": depth,
                "total_slides": len(slide_images),
                "file_path": str(pptx_path)
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

    def _analyze_all_images(self) -> Dict[str, List[Dict]]:
        """
        Analyze all converted images using VLM in parallel

        Returns:
            Dictionary mapping file_key to analysis results
        """
        all_results = {}

        for file_key, image_paths in self.image_registry.items():
            if not image_paths:
                logger.warning(f"No images found for {file_key}")
                all_results[file_key] = []
                continue

            file_info = self.file_hierarchy.get(file_key, {})
            filename = file_info.get("filename", file_key)

            logger.info(f"\nAnalyzing {len(image_paths)} slides from: {filename}")

            try:
                # Analyze slides in parallel
                results = analyze_slides(
                    image_paths,
                    use_parallel=True,
                    max_workers=self.max_vlm_workers,
                    show_progress=True
                )

                all_results[file_key] = results
                logger.info(f"Completed analysis for: {filename}")

            except Exception as e:
                logger.error(f"Error analyzing {filename}: {e}", exc_info=True)
                all_results[file_key] = []

        return all_results

    def _generate_output_files(self, all_results: Dict[str, List[Dict]]) -> None:
        """
        Generate individual JSON output files for each PowerPoint file

        Args:
            all_results: Dictionary mapping file_key to analysis results
        """
        for file_key, slide_results in tqdm(all_results.items(), desc="Generating JSON files"):
            file_info = self.file_hierarchy.get(file_key, {})

            output_data = {
                "file_info": {
                    "filename": file_info.get("filename", file_key),
                    "parent_file": file_info.get("parent_file"),
                    "depth": file_info.get("depth", 0),
                    "total_slides": len(slide_results),
                    "has_embedded_files": len(file_info.get("embedded_files", [])) > 0,
                    "embedded_files": file_info.get("embedded_files", []),
                    "processed_at": datetime.now().isoformat()
                },
                "slides": slide_results
            }

            # Save to individual JSON file
            output_filename = f"{file_key}_analysis.json"
            output_path = self.output_dir / output_filename

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, ensure_ascii=False, indent=2)

            logger.debug(f"Saved: {output_filename}")

    def cleanup_temp_files(self) -> None:
        """Clean up temporary files"""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            logger.info("Cleaned up temporary files")


def main():
    """Main entry point"""
    logger.info("="*80)
    logger.info("VLM Recursive OCR Processor - 2-Step Architecture")
    logger.info("="*80)
    logger.info(f"Data directory: {DATA_DIR}")
    logger.info(f"Output directory: {OUTPUT_DIR}")
    logger.info(f"Temp directory: {TEMP_DIR}")
    logger.info("="*80)

    processor = TwoStepPPTXProcessor(max_vlm_workers=5)

    try:
        # Process all PowerPoint files
        results = processor.process_directory(DATA_DIR)

        logger.info("\n" + "="*80)
        logger.info("Processing completed successfully!")
        logger.info(f"Total files processed: {len(processor.image_registry)}")
        logger.info(f"Total images analyzed: {sum(len(imgs) for imgs in processor.image_registry.values())}")
        logger.info(f"Output saved to: {OUTPUT_DIR}")
        logger.info("="*80)

    except Exception as e:
        logger.error(f"Fatal error during processing: {e}", exc_info=True)
        raise

    finally:
        # Optional: Clean up temporary files
        # Uncomment the line below to automatically delete temp files
        # processor.cleanup_temp_files()
        logger.info(f"\nTemporary files preserved in: {TEMP_DIR}")
        logger.info("To clean up, run: processor.cleanup_temp_files()")


if __name__ == "__main__":
    main()
