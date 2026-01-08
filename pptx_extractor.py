"""
PowerPoint OLE Object Extractor
Extracts embedded OLE objects (ppt/pptx) from PowerPoint files recursively
"""
import os
import tempfile
import zipfile
import shutil
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


def extract_embedded_pptx_from_zip(pptx_path: Path, output_dir: Path) -> List[Path]:
    """
    Extract embedded files directly from PPTX ZIP structure

    Args:
        pptx_path: Path to the source PowerPoint file
        output_dir: Directory to save extracted files

    Returns:
        List of paths to extracted PowerPoint files
    """
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    extracted_files = []

    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            # Look for embedded files in ppt/embeddings/ directory
            embedded_files = [f for f in zip_ref.namelist() if f.startswith('ppt/embeddings/')]

            logger.info(f"Found {len(embedded_files)} files in ppt/embeddings/")

            for idx, file_path in enumerate(embedded_files, 1):
                try:
                    # Extract the file
                    file_data = zip_ref.read(file_path)

                    # Determine file extension by checking file signature
                    ext = _detect_file_extension(file_data)

                    # Skip if not a PowerPoint file
                    if ext not in ['.ppt', '.pptx', '.pptm']:
                        logger.debug(f"Skipping non-PowerPoint file: {file_path} (detected: {ext})")
                        continue

                    # Generate output filename
                    output_filename = f"embedded_{idx}{ext}"
                    output_path = output_dir / output_filename

                    # Save the file
                    with open(output_path, 'wb') as f:
                        f.write(file_data)

                    extracted_files.append(output_path)
                    logger.info(f"Extracted embedded file: {output_filename}")

                except Exception as e:
                    logger.error(f"Error extracting {file_path}: {e}")

    except zipfile.BadZipFile:
        logger.error(f"Invalid PPTX file (not a valid ZIP): {pptx_path}")
    except Exception as e:
        logger.error(f"Error reading PPTX as ZIP: {e}")

    return extracted_files


def _detect_file_extension(file_data: bytes) -> str:
    """
    Detect file extension from file signature (magic bytes)

    Args:
        file_data: Binary file data

    Returns:
        File extension
    """
    if len(file_data) < 8:
        return '.bin'

    # Check for ZIP-based formats (PPTX, DOCX, etc.)
    if file_data[:4] == b'PK\x03\x04':
        # Further check for PPTX by looking for specific content
        if b'ppt/' in file_data[:1000] or b'[Content_Types].xml' in file_data[:2000]:
            return '.pptx'
        elif b'word/' in file_data[:1000]:
            return '.docx'
        elif b'xl/' in file_data[:1000]:
            return '.xlsx'
        return '.zip'

    # Check for old Office format (OLE2)
    if file_data[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
        # This is an OLE2 file, likely .ppt or .doc
        # Default to .ppt for PowerPoint context
        return '.ppt'

    return '.bin'


def extract_embedded_pptx(pptx_path: Path, output_dir: Path) -> List[Path]:
    """
    Convenience function to extract all embedded PowerPoint files
    Uses both ZIP extraction and OLE object parsing

    Args:
        pptx_path: Path to the source PowerPoint file
        output_dir: Directory to save extracted files

    Returns:
        List of paths to extracted PowerPoint files
    """
    # Try ZIP extraction first (more reliable for embedded files)
    extracted_files = extract_embedded_pptx_from_zip(pptx_path, output_dir)

    if extracted_files:
        logger.info(f"Successfully extracted {len(extracted_files)} embedded PowerPoint files using ZIP method")
        return extracted_files

    # Fallback to OLE object extraction
    logger.info("No files found with ZIP method, trying OLE extraction...")
    extractor = PPTXExtractor(pptx_path)
    extractor.load_presentation()
    ole_objects = extractor.extract_ole_objects()
    return extractor.save_ole_objects(ole_objects, output_dir)
