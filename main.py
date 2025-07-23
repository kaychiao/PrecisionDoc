#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Main program entry file
"""

import os
import argparse
import logging
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    """Main program entry point"""
    parser = argparse.ArgumentParser(description="Process PDF files with AI")
    parser.add_argument("--folder", required=True, help="Folder containing PDF files")
    parser.add_argument("--api-key", help="API key for OpenAI or Qwen")
    parser.add_argument("--use-qwen", action="store_true", help="Use Qwen API instead of OpenAI")
    parser.add_argument("--output-folder", default="./output", help="Output folder for results")
    
    args = parser.parse_args()
    
    # Load environment variables from .env file
    load_dotenv()
    
    # Import processor (using new import path)
    from precisiondoc.pdf.pdf_processor import PDFProcessor
    
    processor = PDFProcessor(
        folder_path=args.folder,
        api_key=args.api_key,
        use_qwen=args.use_qwen,
        output_folder=args.output_folder
    )
    
    # results = processor.process_all()
    results = {} # skip this step
    processor.save_consolidated_results(results)

if __name__ == "__main__":
    main()
