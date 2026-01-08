"""
PowerPoint OLE Object Extractor
Extracts embedded OLE objects (ppt/pptx) from PowerPoint files recursively
"""
import os
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple
from pptx import Presentation
from pptx.shapes.graphfrm import GraphicFrame
from pptx.oxml import parse_xml
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PPTXExtractor:
    """Extract OLE objects from PowerPoint files"""

    def __init__(self, pptx_path: Path):
        """
        Initialize the extractor

        Args:
            pptx_path: Path to the PowerPoint file
        """
        self.pptx_path = Path(pptx_path)
        self.presentation = None

    def load_presentation(self) -> None:
        """Load the PowerPoint presentation"""
        try:
            self.presentation = Presentation(str(self.pptx_path))
            logger.info(f"Loaded presentation: {self.pptx_path.name}")
        except Exception as e:
            logger.error(f"Failed to load presentation: {e}")
            raise

    def extract_ole_objects(self) -> List[Dict]:
        """
        Extract all OLE objects from the presentation

        Returns:
            List of dictionaries containing OLE object information
        """
        if not self.presentation:
            self.load_presentation()

        ole_objects = []

        for slide_idx, slide in enumerate(self.presentation.slides, 1):
            logger.info(f"Processing slide {slide_idx}/{len(self.presentation.slides)}")

            for shape_idx, shape in enumerate(slide.shapes):
                # Check if shape contains OLE object
                if self._is_ole_object(shape):
                    ole_info = self._extract_ole_data(shape, slide_idx, shape_idx)
                    if ole_info:
                        ole_objects.append(ole_info)

        logger.info(f"Found {len(ole_objects)} OLE objects")
        return ole_objects

    def _is_ole_object(self, shape) -> bool:
        """
        Check if a shape is an OLE object

        Args:
            shape: PowerPoint shape object

        Returns:
            True if shape is an OLE object
        """
        try:
            # Check if shape has oleObject element
            if hasattr(shape, '_element'):
                xml = shape._element.xml
                return b'oleObject' in xml or b'embed' in xml
            return False
        except Exception as e:
            logger.debug(f"Error checking OLE object: {e}")
            return False

    def _extract_ole_data(self, shape, slide_idx: int, shape_idx: int) -> Dict:
        """
        Extract OLE object data from a shape

        Args:
            shape: PowerPoint shape containing OLE object
            slide_idx: Slide index
            shape_idx: Shape index within the slide

        Returns:
            Dictionary with OLE object information
        """
        try:
            # Get the relationship ID for the embedded object
            ole_elem = shape._element

            # Try to find the embedded object relationship
            for rel in shape.part.rels.values():
                if 'oleObject' in rel.reltype or 'package' in rel.reltype:
                    # Extract the blob data
                    blob_data = rel.target_part.blob

                    # Determine file extension
                    content_type = rel.target_part.content_type
                    ext = self._get_extension_from_content_type(content_type)

                    return {
                        'slide_idx': slide_idx,
                        'shape_idx': shape_idx,
                        'blob_data': blob_data,
                        'extension': ext,
                        'content_type': content_type,
                        'filename': f"ole_s{slide_idx}_sh{shape_idx}{ext}"
                    }

            return None

        except Exception as e:
            logger.error(f"Error extracting OLE data from slide {slide_idx}, shape {shape_idx}: {e}")
            return None

    def _get_extension_from_content_type(self, content_type: str) -> str:
        """
        Get file extension from content type

        Args:
            content_type: MIME type or content type string

        Returns:
            File extension with dot (e.g., '.pptx')
        """
        content_type_map = {
            'application/vnd.ms-powerpoint': '.ppt',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
            'application/vnd.ms-powerpoint.presentation.macroEnabled.12': '.pptm',
        }

        return content_type_map.get(content_type, '.bin')

    def save_ole_objects(self, ole_objects: List[Dict], output_dir: Path) -> List[Path]:
        """
        Save extracted OLE objects to files

        Args:
            ole_objects: List of OLE object dictionaries
            output_dir: Directory to save extracted files

        Returns:
            List of paths to saved files
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        saved_files = []

        for ole_obj in ole_objects:
            try:
                # Only save PowerPoint files
                if ole_obj['extension'] in ['.ppt', '.pptx', '.pptm']:
                    output_path = output_dir / ole_obj['filename']

                    with open(output_path, 'wb') as f:
                        f.write(ole_obj['blob_data'])

                    saved_files.append(output_path)
                    logger.info(f"Saved OLE object to: {output_path}")

            except Exception as e:
                logger.error(f"Error saving OLE object {ole_obj['filename']}: {e}")

        return saved_files


def extract_embedded_pptx(pptx_path: Path, output_dir: Path) -> List[Path]:
    """
    Convenience function to extract all embedded PowerPoint files

    Args:
        pptx_path: Path to the source PowerPoint file
        output_dir: Directory to save extracted files

    Returns:
        List of paths to extracted PowerPoint files
    """
    extractor = PPTXExtractor(pptx_path)
    extractor.load_presentation()
    ole_objects = extractor.extract_ole_objects()
    return extractor.save_ole_objects(ole_objects, output_dir)
