"""
VLM (Vision Language Model) Analyzer
Analyzes slide images using OpenAI's GPT-4o Vision API with retry logic and parallel processing
"""
import base64
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import AzureOpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from tqdm import tqdm
from config import AZURE_ENDPOINT, AZURE_API_KEY, AZURE_API_VERSION, GPT_MODEL, SYSTEM_PROMPT, JSON_SCHEMA

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class VLMAnalyzer:
    """Analyze slide images using GPT-4o Vision with error recovery and parallel processing"""

    def __init__(
        self,
        api_key: str = None,
        model: str = None,
        max_workers: int = 5,
        max_retries: int = 3
    ):
        """
        Initialize the VLM analyzer

        Args:
            api_key: Azure OpenAI API key (defaults to config.AZURE_API_KEY)
            model: Model name (defaults to config.GPT_MODEL)
            max_workers: Maximum number of parallel workers for analysis
            max_retries: Maximum number of retry attempts for failed requests
        """
        self.api_key = api_key or AZURE_API_KEY
        self.model = model or GPT_MODEL
        self.client = AzureOpenAI(
            azure_endpoint=AZURE_ENDPOINT,
            api_key=self.api_key,
            api_version=AZURE_API_VERSION
        )
        self.max_workers = max_workers
        self.max_retries = max_retries

    def encode_image(self, image_path: Path) -> str:
        """
        Encode image to base64 string

        Args:
            image_path: Path to image file

        Returns:
            Base64 encoded image string
        """
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((Exception,)),
        reraise=True
    )
    def _call_vision_api(self, base64_image: str, user_prompt: str) -> str:
        """
        Call GPT-4o Vision API with retry logic

        Args:
            base64_image: Base64 encoded image
            user_prompt: User prompt text

        Returns:
            API response text
        """
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": SYSTEM_PROMPT
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": user_prompt
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=2000,
            temperature=0.1
        )
        return response.choices[0].message.content.strip()

    def analyze_slide_image(
        self,
        image_path: Path,
        slide_number: Optional[int] = None
    ) -> Dict:
        """
        Analyze a single slide image using GPT-4o Vision

        Args:
            image_path: Path to slide image
            slide_number: Optional slide number for tracking

        Returns:
            Dictionary containing extracted information
        """
        try:
            # Encode image
            base64_image = self.encode_image(image_path)

            # Prepare the prompt with JSON schema
            user_prompt = f"""
Analyze this PowerPoint slide image and extract the following information as a JSON object:

{json.dumps(JSON_SCHEMA, indent=2, ensure_ascii=False)}

Instructions:
- For 'visual_references', provide a list of descriptions for any charts, diagrams, tables, or visual elements
- For 'confidence_scores', provide a confidence level (0.0 to 1.0) for each field you extracted
- If a field is not present in the slide, use an empty string or empty list
- Focus on technical problem-cause-solution analysis

Return ONLY the JSON object, no additional text.
"""

            # Call API with retry logic
            result_text = self._call_vision_api(base64_image, user_prompt)

            # Parse JSON response
            result_json = self._parse_json_response(result_text)

            # Add metadata
            if slide_number is not None:
                result_json['slide_number'] = slide_number
            result_json['image_filename'] = image_path.name

            logger.debug(f"Successfully analyzed: {image_path.name}")
            return result_json

        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON response for {image_path.name}: {e}")
            return self._create_error_response(
                f"JSON parsing error: {e}",
                image_path.name,
                slide_number
            )

        except Exception as e:
            logger.error(f"Error analyzing slide image {image_path.name}: {e}")
            return self._create_error_response(
                str(e),
                image_path.name,
                slide_number
            )

    def _parse_json_response(self, result_text: str) -> Dict:
        """
        Parse JSON response from API, handling markdown code blocks

        Args:
            result_text: Raw API response text

        Returns:
            Parsed JSON dictionary
        """
        # Remove markdown code blocks if present
        if result_text.startswith("```json"):
            result_text = result_text[7:]
        if result_text.startswith("```"):
            result_text = result_text[3:]
        if result_text.endswith("```"):
            result_text = result_text[:-3]

        return json.loads(result_text.strip())

    def analyze_multiple_slides(
        self,
        image_paths: List[Path],
        use_parallel: bool = True,
        show_progress: bool = True
    ) -> List[Dict]:
        """
        Analyze multiple slide images with optional parallel processing

        Args:
            image_paths: List of paths to slide images
            use_parallel: Whether to use parallel processing
            show_progress: Whether to show progress bar

        Returns:
            List of dictionaries containing extracted information
        """
        if use_parallel:
            return self._analyze_parallel(image_paths, show_progress)
        else:
            return self._analyze_sequential(image_paths, show_progress)

    def _analyze_sequential(
        self,
        image_paths: List[Path],
        show_progress: bool
    ) -> List[Dict]:
        """
        Analyze slides sequentially

        Args:
            image_paths: List of paths to slide images
            show_progress: Whether to show progress bar

        Returns:
            List of analysis results
        """
        results = []
        iterator = tqdm(image_paths, desc="Analyzing slides") if show_progress else image_paths

        for idx, image_path in enumerate(iterator, 1):
            if not show_progress:
                logger.info(f"Analyzing slide {idx}/{len(image_paths)}: {image_path.name}")

            result = self.analyze_slide_image(image_path, slide_number=idx)
            results.append(result)

        logger.info(f"Completed analysis of {len(results)} slides")
        return results

    def _analyze_parallel(
        self,
        image_paths: List[Path],
        show_progress: bool
    ) -> List[Dict]:
        """
        Analyze slides in parallel using ThreadPoolExecutor

        Args:
            image_paths: List of paths to slide images
            show_progress: Whether to show progress bar

        Returns:
            List of analysis results (ordered by slide number)
        """
        results = {}

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_idx = {
                executor.submit(
                    self.analyze_slide_image,
                    image_path,
                    idx
                ): idx
                for idx, image_path in enumerate(image_paths, 1)
            }

            # Process completed tasks with progress bar
            iterator = as_completed(future_to_idx)
            if show_progress:
                iterator = tqdm(iterator, total=len(image_paths), desc="Analyzing slides")

            for future in iterator:
                idx = future_to_idx[future]
                try:
                    result = future.result()
                    results[idx] = result
                except Exception as e:
                    logger.error(f"Error in parallel processing for slide {idx}: {e}")
                    results[idx] = self._create_error_response(
                        str(e),
                        f"slide_{idx}",
                        idx
                    )

        # Return results ordered by slide number
        ordered_results = [results[i] for i in sorted(results.keys())]
        logger.info(f"Completed parallel analysis of {len(ordered_results)} slides")
        return ordered_results

    def _create_error_response(
        self,
        error_message: str,
        image_filename: str = "",
        slide_number: Optional[int] = None
    ) -> Dict:
        """
        Create an error response following the JSON schema

        Args:
            error_message: Error message
            image_filename: Name of the image file
            slide_number: Optional slide number

        Returns:
            Dictionary with error information
        """
        error_dict = {
            "title": "",
            "problem_symptom": "",
            "cause": "",
            "countermeasure": "",
            "summary": "",
            "visual_references": [],
            "additional_notes": f"Error during analysis: {error_message}",
            "confidence_scores": {},
            "error": error_message,
            "image_filename": image_filename
        }

        if slide_number is not None:
            error_dict["slide_number"] = slide_number

        return error_dict


def analyze_slides(
    image_paths: List[Path],
    use_parallel: bool = True,
    max_workers: int = 5,
    show_progress: bool = True
) -> List[Dict]:
    """
    Convenience function to analyze slide images

    Args:
        image_paths: List of paths to slide images
        use_parallel: Whether to use parallel processing
        max_workers: Maximum number of parallel workers
        show_progress: Whether to show progress bar

    Returns:
        List of analysis results
    """
    analyzer = VLMAnalyzer(max_workers=max_workers)
    return analyzer.analyze_multiple_slides(
        image_paths,
        use_parallel=use_parallel,
        show_progress=show_progress
    )
