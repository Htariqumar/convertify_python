# pdf_to_word/__init__.py
"""
pdf_to_word — modular PDF → DOCX conversion pipeline.

Public surface:
    from pdf_to_word import convert, ConversionResult, InvalidFileError
"""
from .pipeline import convert, ConversionResult, InvalidFileError

__all__ = ["convert", "ConversionResult", "InvalidFileError"]
