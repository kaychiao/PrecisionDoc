"""
Word document processing utilities package.
This package provides utilities for working with Word documents.
"""

from precisiondoc.utils.word.document_formatting import DocumentFormatter
from precisiondoc.utils.word.table_utils import TableUtils
from precisiondoc.utils.word.content_formatting import ContentFormatter
from precisiondoc.utils.word.evidence_processing import EvidenceProcessor
from precisiondoc.utils.word.image_utils import ImageUtils
from precisiondoc.utils.word.export_utils import ExportUtils
from precisiondoc.utils.log_utils import setup_logger

import pandas as pd
import os

# For backward compatibility, create a WordUtils class that combines functionality from all other classes
class WordUtils(DocumentFormatter, TableUtils, ContentFormatter, 
               EvidenceProcessor, ImageUtils, ExportUtils):
    """Word document processing utility class - Facade for all Word utilities"""
    
    @staticmethod
    def export_evidence_to_word(excel_file, word_file, output_folder=None, multi_line_text=True, show_borders=True, exclude_columns=None, page_settings=None):
        """
        Export precision evidence from Excel to Word document.
        This method provides backward compatibility with the original function signature.
        
        Args:
            excel_file: Path to Excel file or DataFrame
            word_file: Path to output Word file
            output_folder: Output folder path, used to find images (legacy parameter)
            multi_line_text: If True, split text by newlines in the left cell
            show_borders: If True, show table borders
            exclude_columns: Columns to exclude from evidence text
            page_settings: Dictionary with page settings for Word document.
                          Supported keys: 'orientation' ('portrait' or 'landscape'),
                          'margins' (dict with 'left', 'right', 'top', 'bottom' in inches)
            
        Returns:
            str: Path to the saved Word document
        """
        logger = setup_logger(__name__)
        
        # Check if excel_file is already a DataFrame
        if isinstance(excel_file, pd.DataFrame):
            df = excel_file
        else:
            # Read Excel file into DataFrame
            df = pd.read_excel(excel_file)
            
        # Filter precision evidence if the column exists
        if 'is_precision_evidence' in df.columns:
            # Handle different types of is_precision_evidence values
            # Possible values: True, 'True', 'true', 1, '1', 'yes', 'Y', etc.
            true_values = [True, 'True', 'true', 1, '1', 'yes', 'Y']
            evidence_df = df[df['is_precision_evidence'].isin(true_values)].copy()
            
            logger.info(f"Filtering precision evidence: found {len(evidence_df)} rows out of {len(df)} total rows")
            
            if evidence_df.empty:
                logger.warning("No precision evidence found in Excel file")
                return word_file
                
            # Call the new implementation with the DataFrame
            return ExportUtils.export_evidence_to_word(evidence_df, word_file, 
                                                     output_folder=output_folder,
                                                     multi_line_text=multi_line_text, 
                                                     show_borders=show_borders,
                                                     exclude_columns=exclude_columns,
                                                     page_settings=page_settings)
        else:
            logger.info("No 'is_precision_evidence' column found, using all rows")
            # If no filtering needed, use the entire DataFrame
            return ExportUtils.export_evidence_to_word(df, word_file, 
                                                     output_folder=output_folder,
                                                     multi_line_text=multi_line_text,
                                                     show_borders=show_borders,
                                                     exclude_columns=exclude_columns,
                                                     page_settings=page_settings)

# Export main classes to make them directly importable from precisiondoc.utils.word
__all__ = [
    'WordUtils',
    'DocumentFormatter',
    'TableUtils',
    'ContentFormatter',
    'EvidenceProcessor',
    'ImageUtils',
    'ExportUtils'
]
