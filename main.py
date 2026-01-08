"""
VLM Recursive OCR - Main Script (Unified Execution)
Runs both Step 1 (conversion) and Step 2 (analysis) sequentially

For large datasets (10,000+ slides), consider running steps separately:
  python step1_convert.py
  python step2_analyze.py --workers 10
"""
import logging
import argparse
from pathlib import Path

from config import DATA_DIR, OUTPUT_DIR
from step1_convert import Step1Converter
from step2_analyze import Step2Analyzer

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def main():
    """
    Main entry point - runs both steps sequentially
    """
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="VLM Recursive OCR - Complete Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Run complete pipeline (both steps)
  python main.py

  # Run with custom VLM worker count
  python main.py --workers 10

  # Run steps separately (recommended for large datasets)
  python step1_convert.py
  python step2_analyze.py --workers 10
        """
    )
    parser.add_argument(
        '--workers',
        type=int,
        default=5,
        help='Number of parallel workers for VLM analysis (default: 5)'
    )
    parser.add_argument(
        '--skip-step1',
        action='store_true',
        help='Skip Step 1 (conversion) and only run Step 2 (analysis)'
    )
    parser.add_argument(
        '--skip-step2',
        action='store_true',
        help='Skip Step 2 (analysis) and only run Step 1 (conversion)'
    )
    args = parser.parse_args()

    logger.info("="*80)
    logger.info("VLM Recursive OCR - Complete Pipeline")
    logger.info("="*80)
    logger.info(f"Data directory: {DATA_DIR}")
    logger.info(f"Output directory: {OUTPUT_DIR}")
    logger.info(f"VLM workers: {args.workers}")
    logger.info("="*80)

    try:
        # Step 1: Image Conversion
        if not args.skip_step1:
            logger.info("\n" + "="*80)
            logger.info("Starting STEP 1: Image Conversion")
            logger.info("="*80)

            converter = Step1Converter()
            metadata = converter.process_all()

            logger.info("\nStep 1 completed successfully!")
        else:
            logger.info("\nSkipping Step 1 (--skip-step1)")

        # Step 2: VLM Analysis
        if not args.skip_step2:
            logger.info("\n" + "="*80)
            logger.info("Starting STEP 2: VLM Analysis")
            logger.info("="*80)

            analyzer = Step2Analyzer(max_vlm_workers=args.workers)
            summary = analyzer.process_all()

            logger.info("\nStep 2 completed successfully!")
        else:
            logger.info("\nSkipping Step 2 (--skip-step2)")

        # Final summary
        if not args.skip_step1 and not args.skip_step2:
            logger.info("\n" + "="*80)
            logger.info("ALL PROCESSING COMPLETED!")
            logger.info("="*80)
            logger.info(f"Results saved to: {OUTPUT_DIR}")
            logger.info("="*80)

    except FileNotFoundError as e:
        logger.error(str(e))
        raise

    except Exception as e:
        logger.error(f"Fatal error during processing: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
