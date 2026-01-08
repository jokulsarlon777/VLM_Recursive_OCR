"""
PowerPoint Slide to Image Converter
Converts PowerPoint slides to images using COM automation (Windows only)
"""
import os
import logging
from pathlib import Path
from typing import List
import win32com.client
import pythoncom
from tqdm import tqdm

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class SlideImageConverter:
    """Convert PowerPoint slides to images using COM automation"""

    def __init__(self):
        """Initialize the converter"""
        self.powerpoint = None

    def __enter__(self):
        """Context manager entry"""
        self._initialize_powerpoint()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self._cleanup_powerpoint()

    def _initialize_powerpoint(self) -> None:
        """Initialize PowerPoint application via COM"""
        try:
            pythoncom.CoInitialize()
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            self.powerpoint.Visible = False
            logger.info("PowerPoint application initialized")
        except Exception as e:
            logger.error(f"Failed to initialize PowerPoint: {e}")
            raise RuntimeError(
                "Failed to initialize PowerPoint. Make sure Microsoft PowerPoint is installed."
            ) from e

    def _cleanup_powerpoint(self) -> None:
        """Clean up PowerPoint application"""
        try:
            if self.powerpoint:
                self.powerpoint.Quit()
                self.powerpoint = None
            pythoncom.CoUninitialize()
            logger.info("PowerPoint application closed")
        except Exception as e:
            logger.error(f"Error during PowerPoint cleanup: {e}")

    def convert_slides_to_images(
        self,
        pptx_path: Path,
        output_dir: Path,
        image_format: str = "PNG",
        width: int = 1920,
        height: int = 1080,
        show_progress: bool = True
    ) -> List[Path]:
        """
        Convert all slides in a PowerPoint file to images

        Args:
            pptx_path: Path to PowerPoint file
            output_dir: Directory to save slide images
            image_format: Image format (PNG, JPG, BMP)
            width: Image width in pixels
            height: Image height in pixels
            show_progress: Whether to show progress bar

        Returns:
            List of paths to generated slide images
        """
        pptx_path = Path(pptx_path).resolve()
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        if not pptx_path.exists():
            raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")

        image_paths = []

        try:
            # Open presentation
            presentation = self.powerpoint.Presentations.Open(
                str(pptx_path),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False
            )

            logger.info(f"Opened presentation: {pptx_path.name}")
            logger.info(f"Total slides: {presentation.Slides.Count}")

            # Export each slide with progress bar
            slide_range = range(1, presentation.Slides.Count + 1)
            iterator = tqdm(slide_range, desc=f"Converting {pptx_path.name}") if show_progress else slide_range

            for slide_idx in iterator:
                slide = presentation.Slides(slide_idx)

                # Generate output filename
                base_name = pptx_path.stem
                output_filename = f"{base_name}_slide_{slide_idx:03d}.{image_format.lower()}"
                output_path = output_dir / output_filename

                # Export slide as image
                slide.Export(
                    str(output_path),
                    image_format,
                    width,
                    height
                )

                image_paths.append(output_path)
                if not show_progress:
                    logger.info(f"Exported slide {slide_idx} to {output_filename}")

            # Close presentation
            presentation.Close()
            logger.info(f"Successfully converted {len(image_paths)} slides to images")

        except Exception as e:
            logger.error(f"Error converting slides to images: {e}")
            raise

        return image_paths


def convert_pptx_to_images(
    pptx_path: Path,
    output_dir: Path,
    image_format: str = "PNG",
    width: int = 1920,
    height: int = 1080,
    show_progress: bool = True
) -> List[Path]:
    """
    Convenience function to convert PowerPoint slides to images

    Args:
        pptx_path: Path to PowerPoint file
        output_dir: Directory to save slide images
        image_format: Image format (PNG, JPG, BMP)
        width: Image width in pixels
        height: Image height in pixels
        show_progress: Whether to show progress bar

    Returns:
        List of paths to generated slide images
    """
    with SlideImageConverter() as converter:
        return converter.convert_slides_to_images(
            pptx_path, output_dir, image_format, width, height, show_progress
        )
