import os
import re
import glob
import logging
from typing import Dict, List
from collections import defaultdict
import fitz  # PyMuPDF

# Configure logging
logger = logging.getLogger(__name__)

def normalize_filename(filename: str) -> str:
    """
    Normalize a filename by replacing spaces and special characters with underscores.
    
    Args:
        filename: Original filename
        
    Returns:
        Normalized filename
    """
    # Replace spaces and special characters with underscores
    normalized = re.sub(r'[^\w\-\.]', '_', filename)
    return normalized

def find_latest_pdfs(folder_path: str) -> Dict[str, str]:
    """
    Find the latest version of each type of PDF in the folder.
    Recursively searches through all subdirectories.
    
    Args:
        folder_path: Path to the folder containing PDF files
        
    Returns:
        Dictionary mapping document type to the path of its latest version
    """
    # Recursively find all PDF files
    pdf_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    logger.info(f"Found {len(pdf_files)} PDF files in {folder_path} and its subdirectories")

    # Group files by document type
    document_groups = defaultdict(list)
    
    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        # Extract document type and year using regex
        # Pattern matches: document type + year
        # Example: "CSCO黑色素瘤诊疗指南2024_20241225015714.pdf" -> group="CSCO黑色素瘤诊疗指南", year=2024
        match = re.match(r'(CSCO[^0-9]+)(\d{4}).*\.pdf', filename)
        if match:
            doc_type, year = match.groups()
            document_groups[doc_type].append((int(year), pdf_path))
    
    # Find the latest version of each document type
    latest_pdfs = {}
    for doc_type, versions in document_groups.items():
        # Sort by year (descending)
        versions.sort(reverse=True)
        latest_year, latest_path = versions[0]
        latest_pdfs[doc_type] = latest_path
        logger.info(f"Latest version of {doc_type}: {os.path.basename(latest_path)} (Year: {latest_year})")
    
    return latest_pdfs

def split_pdf(pdf_path: str, output_folder: str) -> List[str]:
    """
    Split a PDF into individual pages and save each page as a separate PDF.
    If the PDF has already been split, returns the existing page paths.
    
    Args:
        pdf_path: Path to the PDF file
        output_folder: Folder to save the split pages
        
    Returns:
        List of paths to the individual page PDFs
    """
    pdf_filename = os.path.basename(pdf_path)
    base_name = os.path.splitext(pdf_filename)[0]
    
    # Normalize the base name for folder creation
    normalized_base_name = normalize_filename(base_name)
    
    # Create a subfolder for this specific PDF
    pdf_subfolder = os.path.join(output_folder, normalized_base_name)
    os.makedirs(pdf_subfolder, exist_ok=True)
    
    # Check if this PDF has already been split
    existing_pages = glob.glob(os.path.join(pdf_subfolder, f"{base_name}_page_*.pdf"))
    if existing_pages:
        logger.info(f"PDF {pdf_filename} has already been split. Found {len(existing_pages)} existing pages.")
        # Sort by page number
        existing_pages.sort()
        return existing_pages
    
    # If not already split, perform splitting
    # Open the PDF
    doc = fitz.open(pdf_path)
    page_paths = []
    
    # Process each page
    for page_num in range(len(doc)):
        # Create a new PDF with just this page
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
        
        # Save the page with three-digit page number (001, 002, etc.)
        page_path = os.path.join(pdf_subfolder, f"{base_name}_page_{page_num+1:03d}.pdf")
        new_doc.save(page_path)
        new_doc.close()
        page_paths.append(page_path)
        
    doc.close()
    logger.info(f"Split {pdf_filename} into {len(page_paths)} pages")
    return page_paths

def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extract text from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Extracted text from the PDF
    """
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def convert_pdf_to_image(pdf_path: str, output_folder: str = None) -> str:
    """
    Convert a PDF page to an image.
    
    Args:
        pdf_path: Path to the PDF file
        output_folder: Folder to save the image (if None, uses same folder as PDF)
        
    Returns:
        Path to the generated image file
    """
    if output_folder is None:
        output_folder = os.path.dirname(pdf_path)
    
    # Get PDF filename and base name
    pdf_filename = os.path.basename(pdf_path)
    base_name = os.path.splitext(pdf_filename)[0]
    
    # Extract original PDF name (if this is a split page)
    original_pdf_name = base_name
    if "_page_" in base_name:
        original_pdf_name = base_name.split("_page_")[0]
    
    # Normalize the original PDF name for folder creation
    normalized_pdf_name = normalize_filename(original_pdf_name)
    
    # Create subfolder for this PDF's images
    pdf_subfolder = os.path.join(output_folder, normalized_pdf_name)
    os.makedirs(pdf_subfolder, exist_ok=True)
    
    # Generate output image path - use just the page number for the filename
    if "_page_" in base_name:
        page_num = base_name.split("_page_")[1]
        image_path = os.path.join(pdf_subfolder, f"page_{page_num}.png")
    else:
        image_path = os.path.join(pdf_subfolder, f"{base_name}.png")
    
    # Open the PDF
    doc = fitz.open(pdf_path)
    
    # We assume it's a single page PDF (as we're processing split pages)
    if len(doc) > 0:
        page = doc[0]
        
        # Set a higher zoom factor for better resolution
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        
        # Render page to pixmap
        pix = page.get_pixmap(matrix=mat)
        
        # Save pixmap as PNG
        pix.save(image_path)
        
        logger.info(f"Converted PDF to image: {image_path}")
    
    doc.close()
    return image_path
