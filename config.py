"""
Configuration file for VLM Recursive OCR project
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Project paths
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / os.getenv("OUTPUT_DIR", "output")
TEMP_DIR = PROJECT_ROOT / os.getenv("TEMP_DIR", "temp")

# Ensure directories exist
OUTPUT_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)

# OpenAI Configuration
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GPT_MODEL = os.getenv("GPT_MODEL", "gpt-4o")

# Validate API key
if not OPENAI_API_KEY:
    raise ValueError(
        "OPENAI_API_KEY not found in environment variables. "
        "Please create a .env file based on .env.example"
    )

# JSON Schema for output
JSON_SCHEMA = {
    "title": "",
    "problem_symptom": "",
    "cause": "",
    "countermeasure": "",
    "summary": "",
    "visual_references": [],
    "additional_notes": "",
    "confidence_scores": {}
}

# System prompt for GPT-4o
SYSTEM_PROMPT = """You are an expert technical document analyzer specializing in problem-cause-solution documentation.
Analyze the provided PowerPoint slide image and extract information in the following categories:

1. title: Main title or heading of the slide
2. problem_symptom: Description of the problem or symptom being discussed
3. cause: Root cause or reasons for the problem
4. countermeasure: Solutions, countermeasures, or action items to address the problem
5. summary: Brief summary of the entire slide content
6. visual_references: Descriptions of any charts, diagrams, tables, or visual elements (as a list)
7. additional_notes: Any other relevant information not covered above
8. confidence_scores: Your confidence level (0-1) for each extracted field

Return the result as a valid JSON object following this structure."""
