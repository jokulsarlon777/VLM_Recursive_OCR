"""
Step 2: Analyze all converted images using VLM (Parallel)
- Loads metadata from Step 1
- Analyzes all images in parallel using GPT-4o Vision
- Saves analysis results to JSON
"""
import json
import logging
from pathlib import Path
from typing import Dict, List
from datetime import datetime
from tqdm import tqdm

from config import OUTPUT_DIR
from vlm_analyzer import analyze_slides

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class Step2Analyzer:
    """
    Step 2: Analyze all converted images using VLM
    """

    def __init__(
        self,
        output_dir: Path = OUTPUT_DIR,
        max_vlm_workers: int = 5,
        metadata_file: str = "step1_metadata.json"
    ):
        """
        Initialize the analyzer

        Args:
            output_dir: Directory for output JSON files
            max_vlm_workers: Maximum number of parallel workers for VLM analysis
            metadata_file: Metadata file from Step 1
        """
        self.output_dir = Path(output_dir)
        self.max_vlm_workers = max_vlm_workers
        self.metadata_path = self.output_dir / metadata_file

        # Ensure directories exist
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Load metadata from Step 1
        self.file_hierarchy: Dict = {}
        self.image_registry: Dict[str, List[Path]] = {}
        self._load_metadata()

    def _load_metadata(self) -> None:
        """Load metadata from Step 1"""
        if not self.metadata_path.exists():
            raise FileNotFoundError(
                f"Metadata file not found: {self.metadata_path}\n"
                f"Please run 'python step1_convert.py' first."
            )

        with open(self.metadata_path, 'r', encoding='utf-8') as f:
            metadata = json.load(f)

        self.file_hierarchy = metadata.get("file_hierarchy", {})
        # Convert string paths back to Path objects
        image_registry_str = metadata.get("image_registry", {})
        self.image_registry = {
            key: [Path(img) for img in imgs]
            for key, imgs in image_registry_str.items()
        }

        step1_info = metadata.get("step1_info", {})
        total_files = step1_info.get("total_files_processed", 0)
        total_images = step1_info.get("total_images_converted", 0)

        logger.info(f"Loaded metadata from Step 1:")
        logger.info(f"  Total files: {total_files}")
        logger.info(f"  Total images: {total_images}")
        logger.info(f"  Processed at: {step1_info.get('processed_at', 'Unknown')}")

    def process_all(self) -> Dict:
        """
        Main processing pipeline: analyze all images

        Returns:
            Dictionary containing all results
        """
        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 2: Analyzing all images with VLM (parallel processing)")
        logger.info(f"Max parallel workers: {self.max_vlm_workers}")
        logger.info(f"Total files to analyze: {len(self.image_registry)}")
        logger.info(f"{'='*80}\n")

        # Analyze all images
        all_results = self._analyze_all_images()

        # Generate output JSON files
        logger.info(f"\n{'='*80}")
        logger.info(f"Generating JSON output files")
        logger.info(f"{'='*80}\n")

        self._generate_output_files(all_results)

        # Generate summary
        summary = {
            "processing_summary": {
                "total_files_processed": len(self.image_registry),
                "total_images_analyzed": sum(
                    len(results) for results in all_results.values()
                ),
                "max_vlm_workers": self.max_vlm_workers,
                "processed_at": datetime.now().isoformat()
            },
            "file_hierarchy": self.file_hierarchy,
            "results": all_results
        }

        summary_path = self.output_dir / "processing_summary.json"
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

        logger.info(f"\nSaved processing summary to: {summary_path}")

        # Print final summary
        logger.info(f"\n{'='*80}")
        logger.info(f"STEP 2 COMPLETED!")
        logger.info(f"Total files analyzed: {len(self.image_registry)}")
        logger.info(f"Total images analyzed: {sum(len(results) for results in all_results.values())}")
        logger.info(f"Output saved to: {self.output_dir}")
        logger.info(f"{'='*80}\n")

        return summary

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

            # Verify images exist
            existing_images = [img for img in image_paths if Path(img).exists()]
            if not existing_images:
                logger.error(f"No image files found for {file_key}")
                logger.error(f"Expected images in: {image_paths[0].parent if image_paths else 'Unknown'}")
                all_results[file_key] = []
                continue

            if len(existing_images) != len(image_paths):
                logger.warning(
                    f"Some images missing for {file_key}: "
                    f"{len(existing_images)}/{len(image_paths)} found"
                )

            file_info = self.file_hierarchy.get(file_key, {})
            filename = file_info.get("filename", file_key)

            logger.info(f"\nAnalyzing {len(existing_images)} slides from: {filename}")

            try:
                # Analyze slides in parallel
                results = analyze_slides(
                    existing_images,
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


def main():
    """Main entry point for Step 2"""
    logger.info("="*80)
    logger.info("VLM Recursive OCR - STEP 2: VLM Analysis")
    logger.info("="*80)

    # Parse command line arguments for worker count
    import argparse
    parser = argparse.ArgumentParser(description="Step 2: Analyze images with VLM")
    parser.add_argument(
        '--workers',
        type=int,
        default=5,
        help='Number of parallel workers for VLM analysis (default: 5)'
    )
    args = parser.parse_args()

    analyzer = Step2Analyzer(max_vlm_workers=args.workers)

    try:
        summary = analyzer.process_all()

        logger.info("\nAll processing completed!")
        logger.info(f"Check results in: {OUTPUT_DIR}")

    except FileNotFoundError as e:
        logger.error(str(e))
        logger.error("\nPlease run Step 1 first:")
        logger.error("  python step1_convert.py")
        raise

    except Exception as e:
        logger.error(f"Fatal error during Step 2: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
