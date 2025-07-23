import os
import logging
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_LINE_SPACING
from docx.enum.section import WD_ORIENT, WD_SECTION

# Configure logging
logger = logging.getLogger(__name__)

class WordUtils:
    """Word document processing utility class"""
    
    @staticmethod
    def apply_word_document_format(doc):
        """
        Apply uniform formatting to Word document
        
        Args:
            doc: Word document object
        """
        # Set page margins and width
        sections = doc.sections
        for section in sections:
            section.page_width = Cm(21)  # A4 width
            section.page_height = Cm(29.7)  # A4 height
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
        
        # Add heading
        heading = doc.add_heading('Precision Evidence Report', 0)
        heading_format = heading.paragraph_format
        heading_format.space_before = Pt(12)
        heading_format.space_after = Pt(24)
        heading_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
        # Set heading font
        for run in heading.runs:
            run.font.name = 'Microsoft YaHei'
            run.font.size = Pt(16)
            run.font.bold = True
    
    @staticmethod
    def apply_paragraph_format(paragraph):
        """
        Apply paragraph formatting
        
        Args:
            paragraph: Paragraph object
        """
        p_format = paragraph.paragraph_format
        p_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5 line spacing
        p_format.space_before = Pt(12)  # Space before paragraph
        p_format.space_after = Pt(12)   # Space after paragraph
    
    @staticmethod
    def apply_run_format(run):
        """
        Apply formatting to text run object
        
        Args:
            run: Text run object
        """
        run.font.name = 'Microsoft YaHei'  # Set font to Microsoft YaHei
        run.font.size = Pt(11)     # Set font size
    
    @staticmethod
    def apply_separator_format(paragraph):
        """
        Apply formatting to separator paragraph
        
        Args:
            paragraph: Paragraph object
        """
        separator_format = paragraph.paragraph_format
        separator_format.space_before = Pt(12)
        separator_format.space_after = Pt(12)
        
        # Set separator line font
        for run in paragraph.runs:
            run.font.name = 'Microsoft YaHei'
    
    @staticmethod
    def export_evidence_to_word(excel_file, word_file, output_folder):
        """
        Export precision evidence from Excel to Word document
        
        Args:
            excel_file: Excel file path
            word_file: Word file path
            output_folder: Output folder path, used to find images
        """
        try:
            # Import necessary libraries
            import pandas as pd
            
            logger.info(f"Exporting evidence from {excel_file} to {word_file}")
            
            # Read Excel file
            df = pd.read_excel(excel_file)
            
            # Filter precision evidence
            if 'is_precision_evidence' in df.columns:
                evidence_df = df[df['is_precision_evidence'] == True].copy()
                
                if evidence_df.empty:
                    logger.warning("No precision evidence found in Excel file")
                    return
                
                # Create Word document
                doc = Document()
                
                # Apply document formatting
                WordUtils.apply_word_document_format(doc)
                
                # Process evidence row by row
                for idx, row in evidence_df.iterrows():
                    # Create paragraph and apply formatting
                    p = doc.add_paragraph()
                    WordUtils.apply_paragraph_format(p)
                    
                    # Add row data (excluding text column)
                    for col in row.index:
                        if col not in ('text', 'image_path', 'page_number', 'success', 'document_type'):
                            # Check if value is empty, if so display "N/A"
                            value = row[col]
                            if pd.isna(value) or value == '':
                                display_value = "N/A"
                            else:
                                display_value = value
                            run = p.add_run(f"{col}: {display_value}\n")
                            WordUtils.apply_run_format(run)
                    
                    # Add separator line
                    separator = doc.add_paragraph('-' * 50)
                    WordUtils.apply_separator_format(separator)
                    
                    # If image_path exists, add the image directly
                    if 'image_path' in row and not pd.isna(row['image_path']):
                        img_path = row['image_path']
                        if os.path.exists(img_path):
                            try:
                                # Add image and set width, maintaining aspect ratio
                                img_p = doc.add_paragraph()
                                img_p.alignment = 1  # Center alignment
                                img_run = img_p.add_run()
                                img_run.add_picture(img_path, width=Inches(6))
                            except Exception as e:
                                logger.warning(f"Failed to add image {img_path}: {str(e)}")
                        else:
                            logger.warning(f"Image path not found: {img_path}")
                    # Fallback to old method if image_path doesn't exist but page_number does
                    elif 'page_number' in row and not pd.isna(row['page_number']):
                        page_num = int(row['page_number'])
                        doc_type = row.get('document_type', '')
                        
                        # Build possible image paths
                        img_folder = os.path.join(output_folder, 'images')
                        possible_paths = [
                            os.path.join(img_folder, f"{doc_type}_page_{page_num}.png"),
                            os.path.join(img_folder, f"page_{page_num}.png"),
                            os.path.join(output_folder, f"{doc_type}_page_{page_num}.png"),
                            os.path.join(output_folder, f"page_{page_num}.png")
                        ]
                        
                        # Find and add image
                        img_found = False
                        for img_path in possible_paths:
                            if os.path.exists(img_path):
                                try:
                                    # Add image and set width, maintaining aspect ratio
                                    img_p = doc.add_paragraph()
                                    img_p.alignment = 1  # Center alignment
                                    img_run = img_p.add_run()
                                    img_run.add_picture(img_path, width=Inches(6))
                                    img_found = True
                                    break
                                except Exception as e:
                                    logger.warning(f"Failed to add image {img_path}: {str(e)}")
                        
                        if not img_found:
                            logger.warning(f"No image found for page {page_num} of document {doc_type}")
                    
                    # Add page break
                    doc.add_page_break()
                
                # Save Word document
                doc.save(word_file)
                logger.info(f"Evidence exported to Word file: {word_file}")
            else:
                logger.warning("Column 'is_precision_evidence' not found in Excel file")
        except ImportError as ie:
            print(">>> ImportError Details:")
            print(f"Error message: {ie}")
            print(f"Module name: {getattr(ie, 'name', 'Unknown')}")
            print(f"Path: {getattr(ie, 'path', 'Unknown')}")
            import traceback
            print("Stack trace:")
            traceback.print_exc()
            logger.error(f"Required libraries not installed: {ie}")
        except Exception as e:
            logger.error(f"Error exporting evidence to Word: {str(e)}")
