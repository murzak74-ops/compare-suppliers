"""Utility functions for OCR and text preprocessing."""

from __future__ import annotations

import os
import re
from typing import List

from pdf2image import convert_from_path
from PIL import Image
import pytesseract


def convert_pdf_to_images(pdf_path: str) -> List[Image.Image]:
    """Convert a PDF file into a list of PIL images."""
    return convert_from_path(pdf_path)


def perform_ocr(image: Image.Image, lang: str = "rus") -> str:
    """Run OCR on a PIL image using pytesseract."""
    return pytesseract.image_to_string(image, lang=lang)


def clean_text(text: str) -> str:
    """Clean up OCR output.

    This function implements minimal rules that can later be extended.
    """
    # Example: replace common OCR mistakes
    text = text.replace("О", "0").replace("о", "0")
    # Remove non-alphanumeric characters but keep whitespace/newlines
    text = re.sub(r"[^\w\s]", " ", text)
    # Collapse multiple spaces/newlines
    text = re.sub(r"\s+\n", "\n", text)
    text = re.sub(r"\n+", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def extract_text_from_file(file_path: str) -> str:
    """Extract and clean text from an image or PDF file."""
    ext = os.path.splitext(file_path.lower())[1]
    texts: List[str] = []
    if ext == ".pdf":
        images = convert_pdf_to_images(file_path)
        for img in images:
            texts.append(perform_ocr(img))
    else:
        img = Image.open(file_path)
        texts.append(perform_ocr(img))
    raw_text = "\n".join(texts)
    return clean_text(raw_text)
