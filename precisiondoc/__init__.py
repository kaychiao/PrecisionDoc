"""
PrecisionDoc - Document processing and evidence extraction package

This package provides tools for processing PDF documents, extracting evidence,
and generating structured outputs in various formats (JSON, Excel, Word).

Main Features:
-------------
1. PDF Processing: Split PDFs into pages, convert to images, and process with AI
2. Evidence Extraction: Extract structured evidence from document images
3. Export Utilities: Export evidence to JSON, Excel, and Word formats
4. Word Document Generation: Create formatted Word documents with evidence tables

Usage Examples:
--------------
Process a PDF folder:
    >>> from precisiondoc import process_pdf
    >>> results = process_pdf(
    ...     folder_path="path/to/pdfs",
    ...     output_folder="./output",
    ...     ai_settings={"api_key": "your-api-key", "base_url": "https://api.example.com/v1", "model": "gpt-4"}
    ... )

Convert Excel evidence to Word:
    >>> from precisiondoc import excel_to_word
    >>> word_file = excel_to_word(
    ...     excel_file="path/to/evidence.xlsx",
    ...     word_file="path/to/output.docx",  # Optional
    ...     multi_line_text=True,  # Optional
    ...     show_borders=True  # Optional
    ... )
"""

__version__ = '0.1.4'

# Import standard libraries
import os
import tempfile
import shutil
from typing import Dict, List, Optional, Union, Any

# Import core components for API exposure
from .pdf.pdf_processor import PDFProcessor
from .utils.word import WordUtils, ExportUtils
from .utils.data_utils import DataUtils
from .utils.word.image_utils import ImageUtils
from .config.models import AISettings, PageSettings, validate_ai_settings, validate_page_settings

# Convenience functions for common operations
def process_pdf(folder_path: str, output_folder: str = "./output", 
                ai_settings: Optional[Union[Dict[str, Any], AISettings]] = None, 
                **legacy_params) -> Dict:
    """
    Process PDF files in a folder and generate evidence extraction results.
    
    Args:
        folder_path (str): Path to folder containing PDF files
        output_folder (str, optional): Output folder for results. Defaults to "./output".
        ai_settings (dict or AISettings, optional): Dictionary with AI settings or AISettings instance.
                                 Supported keys: 'api_key', 'base_url', 'model'
        **legacy_params: Legacy parameters (api_key, base_url, model) for backward compatibility.
                        These are deprecated since v0.1.4 and will be removed in a future version.
                        Please use ai_settings instead.
        
    Returns:
        dict: Dictionary with processing results
    """
    # Use Pydantic model to validate AI settings
    validated_ai_settings = validate_ai_settings(ai_settings, **legacy_params)
    
    processor = PDFProcessor(
        folder_path=folder_path,
        output_folder=output_folder,
        ai_settings=validated_ai_settings
    )
    results = processor.process_all()
    processor.save_consolidated_results(results)
    return results

def process_single_pdf(pdf_path: str, doc_type: Optional[str] = None, output_folder: str = "./output", 
                      multi_line_text: bool = True, show_borders: bool = True, 
                      exclude_columns: Optional[List[str]] = None, 
                      page_settings: Optional[Union[Dict[str, Any], PageSettings]] = None, 
                      ai_settings: Optional[Union[Dict[str, Any], AISettings]] = None, 
                      **legacy_params) -> Dict:
    """
    Process a single PDF file and generate evidence extraction results.
    
    Args:
        pdf_path (str): Path to the PDF file to process
        doc_type (str, optional): Document type/name. If None, will use the PDF filename without extension.
        output_folder (str, optional): Output folder for results. Defaults to "./output".
        multi_line_text (bool, optional): Whether to use multi-line text format in Word output. Defaults to True.
        show_borders (bool, optional): Whether to show borders in Word tables. Defaults to True.
        exclude_columns (list, optional): List of column names to exclude from Word output.
        page_settings (dict or PageSettings, optional): Dictionary with page settings for Word document or PageSettings instance.
                                   Supported keys: 'orientation' ('portrait' or 'landscape'),
                                   'margins' (dict with 'left', 'right', 'top', 'bottom' in inches).
        ai_settings (dict or AISettings, optional): Dictionary with AI settings or AISettings instance.
                                 Supported keys: 'api_key', 'base_url', 'model'
        **legacy_params: Legacy parameters (api_key, base_url, model) for backward compatibility.
                        These are deprecated since v0.1.4 and will be removed in a future version.
                        Please use ai_settings instead.
        
    Returns:
        dict: Dictionary with processing results for the PDF
    """
    # Use Pydantic model to validate AI settings
    validated_ai_settings = validate_ai_settings(ai_settings, **legacy_params)
    
    # Validate page settings parameters
    validated_page_settings = validate_page_settings(page_settings)
    
    # Create processor with the PDF's parent folder
    processor = PDFProcessor(
        folder_path=os.path.dirname(pdf_path),
        output_folder=output_folder,
        ai_settings=validated_ai_settings,
        page_settings=validated_page_settings
    )
    
    # Store Word formatting parameters in the processor
    processor.multi_line_text = multi_line_text
    processor.show_borders = show_borders
    processor.exclude_columns = exclude_columns
    
    # Process the single PDF file
    results = processor.process_single(pdf_path, doc_type)
    
    return results

def excel_to_word(excel_file, word_file=None, output_folder=None, 
                 multi_line_text=True, show_borders=True, exclude_columns=None,
                 page_settings: Optional[Union[Dict[str, Any], PageSettings]] = None):
    """
    Convert Excel file with evidence data to formatted Word document
    
    Args:
        excel_file (str or DataFrame): Path to Excel file or pandas DataFrame
        word_file (str, optional): Path to output Word file (if None, will be generated from excel_file)
        output_folder (str, optional): Output folder path, used to find images
        multi_line_text (bool, optional): If True, split text by newlines in the left cell. Defaults to True.
        show_borders (bool, optional): If True, show table borders. Defaults to True.
        exclude_columns (list, optional): Columns to exclude from evidence text
        page_settings (dict or PageSettings, optional): Dictionary with page settings for Word document or PageSettings instance.
                                   Supported keys: 'orientation' ('portrait' or 'landscape'),
                                   'margins' (dict with 'left', 'right', 'top', 'bottom' in inches).
        
    Returns:
        str: Path to the generated Word file
    """
    # Validate page settings parameters
    validated_page_settings = validate_page_settings(page_settings)
    
    # If page_settings is a Pydantic model, convert it to a dictionary
    page_settings_dict = None
    if validated_page_settings:
        page_settings_dict = validated_page_settings.model_dump()
    
    return WordUtils.export_evidence_to_word(
        excel_file=excel_file,
        word_file=word_file,
        output_folder=output_folder,
        multi_line_text=multi_line_text,
        show_borders=show_borders,
        exclude_columns=exclude_columns,
        page_settings=page_settings_dict
    )

# Export main classes and functions
__all__ = [
    'PDFProcessor',
    'WordUtils',
    'ExportUtils',
    'DataUtils',
    'ImageUtils',
    'process_pdf',
    'process_single_pdf',
    'excel_to_word',
]
