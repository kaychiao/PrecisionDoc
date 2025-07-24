import os
import json
import re
import pandas as pd
from typing import Dict, List
from dotenv import load_dotenv
from datetime import datetime

# Import utility modules
from ..pdf.pdf_utils import find_latest_pdfs, split_pdf, extract_text_from_pdf, convert_pdf_to_image
from ..ai.ai_client import AIClient
from ..utils.word_utils import WordUtils
from ..utils.data_utils import DataUtils
from ..utils.log_utils import setup_logger

# Setup logger for this module
logger = setup_logger(__name__)

class PDFProcessor:
    """Main class for processing PDF documents"""
    
    def __init__(self, folder_path: str, api_key: str = None, use_qwen: bool = False, output_folder: str = "./output"):
        """
        Initialize the PDF processor.
        
        Args:
            folder_path: Path to the folder containing PDF files
            api_key: API key for OpenAI or Qwen. If None, will try to load from environment variables.
            use_qwen: If True, use Qwen API instead of OpenAI
            output_folder: Path to the output folder for results. Default is "./output"
        """
        self.folder_path = folder_path
        self.use_qwen = use_qwen
        
        # Set output directories
        self.output_folder = output_folder
        self.pages_folder = os.path.join(self.output_folder, "pages")
        
        # Create output directories if they don't exist
        os.makedirs(self.output_folder, exist_ok=True)
        os.makedirs(self.pages_folder, exist_ok=True)
        
        # Initialize AI client
        self.ai_client = AIClient(api_key=api_key, use_qwen=use_qwen)
    
    def process_page_with_ai(self, page_path: str) -> Dict:
        """
        Process a single PDF page with AI.
        Skip processing for table of contents or reference pages.
        
        Args:
            page_path: Path to the PDF page
            
        Returns:
            Dictionary containing the AI's response or a skip message
        """
        # Extract text from PDF page
        page_text = extract_text_from_pdf(page_path)
        
        # Create an images folder if using Qwen
        if self.use_qwen:
            images_folder = os.path.join(self.output_folder, "images")
            os.makedirs(images_folder, exist_ok=True)
            
            # Convert PDF to image
            image_path = convert_pdf_to_image(page_path, images_folder)
            logger.info(f"Converted PDF to image: {image_path}")
        else:
            image_path = None
        
        # Use AI to identify page type
        page_type_result = self.ai_client.identify_page_type(page_text, image_path)
        page_type = page_type_result.get("page_type", "content")
        
        logger.info(f"Identified page type: {page_type} for {os.path.basename(page_path)}")
        
        # Skip processing for non-content pages
        if page_type != "content":
            return {
                "success": True,
                "content": f"This page appears to be a {page_type} page and was skipped for detailed AI processing.",
                "page_type": page_type
            }
        
        # Process content with AI
        if self.use_qwen and image_path:
            # Use image-based processing with Qwen
            logger.info(f"Processing page with Qwen image API: {os.path.basename(page_path)}")
            result = self.ai_client.process_image(image_path)
        else:
            # Use text-based processing
            logger.info(f"Processing page with text API: {os.path.basename(page_path)}")
            result = self.ai_client.process_text(page_text)
        
        # Add page type to result
        if "page_type" not in result:
            result["page_type"] = page_type
            
        # Add image path to result for later use
        if image_path:
            result["image_path"] = image_path
            
        return result
    
    def process_all(self) -> Dict[str, List[Dict]]:
        """
        Process all PDFs in the folder.
        
        Returns:
            Dictionary mapping document type to list of page results
        """
        # Find latest PDFs
        latest_pdfs = find_latest_pdfs(self.folder_path)
        
        if not latest_pdfs:
            logger.warning(f"No PDF files found in {self.folder_path}")
            return {}
        
        results = {}

        # need to remove later
        # latest_pdfs = {k:v for k,v in latest_pdfs.items() if "非小细胞肺癌" in v}
        
        # Create a consolidated output file path that will be used throughout processing
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        consolidated_output_file = os.path.join(self.output_folder, f"all_results_{timestamp}.json")
        
        results = {}


        # Process each PDF
        for doc_type, pdf_path in latest_pdfs.items():
            logger.info(f"Processing {doc_type}: {os.path.basename(pdf_path)}")
            
            # Split PDF into pages
            page_paths = split_pdf(pdf_path, self.pages_folder)
            
            # Process each page with AI
            page_results = []
            for i, page_path in enumerate(page_paths):
                logger.info(f"Processing page {i+1}/{len(page_paths)} of {doc_type}")
                try:
                    result = self.process_page_with_ai(page_path)
                    page_results.append(result)
                    
                    # Save intermediate results after each page (optional, may be too frequent)
                    if (i+1) % 5 == 0:  # Save every 5 pages
                        temp_results = {doc_type: page_results}
                        self.save_consolidated_results(temp_results, consolidated_output_file)
                except Exception as e:
                    logger.error(f"Error processing page {i+1} of {doc_type}: {str(e)}")
                    # Continue with next page despite errors
            
            results[doc_type] = page_results
            
            # Save individual document results to JSON file
            output_file = os.path.join(self.output_folder, f"{doc_type}_results.json")
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(page_results, f, ensure_ascii=False, indent=2)
            logger.info(f"Saved results for {doc_type} to {output_file}")
            
            # Save consolidated results after each document is processed
            self.save_consolidated_results({doc_type: page_results}, consolidated_output_file)
        
        return results
    
    def save_consolidated_results(self, results: Dict[str, List[Dict]], output_file: str = None) -> str:
        """
        Save consolidated results to a JSON file and an Excel file.
        
        Args:
            results: Dictionary mapping document type to list of page results
            output_file: Path to the output file (default: {pdf_name}_results.json in output_folder)
        
        Returns:
            Path to the output file
        """
        # Process each PDF separately
        json_files = []
        
        for pdf_name, page_results in results.items():
            # Normalize PDF name for file naming
            normalized_pdf_name = re.sub(r'[^\w\-\.]', '_', pdf_name)
            
            # Create output file path based on PDF name
            if output_file:
                # If output_file is provided, use its directory but with normalized PDF name
                output_dir = os.path.dirname(output_file)
                base_filename = f"{normalized_pdf_name}_results.json"
                pdf_output_file = os.path.join(output_dir, base_filename)
            else:
                # Otherwise use output_folder with normalized PDF name
                pdf_output_file = os.path.join(self.output_folder, f"{normalized_pdf_name}_results.json")
            
            # Save JSON results for this PDF
            pdf_results = {pdf_name: page_results}
            json_file = DataUtils.handle_json_file(pdf_results, pdf_output_file, self.output_folder)
            json_files.append(json_file)
            
            # Create Excel file path
            excel_file = json_file.replace('.json', '.xlsx')
            
            # Convert nested results to flat structure
            all_rows = DataUtils.convert_to_flat_structure(pdf_results)
            
            # Save to Excel
            DataUtils.save_to_excel(all_rows, excel_file)
            
            # Export evidence to Word
            word_file = excel_file.replace('.xlsx', '_evidence.docx')
            WordUtils.export_evidence_to_word(excel_file, word_file, self.output_folder)
            
            logger.info(f"Saved results for {pdf_name} to {json_file}, {excel_file}, and {word_file}")
        
        # Return the list of JSON files or the first one if there's only one
        if len(json_files) == 1:
            return json_files[0]
        return json_files
