#!/usr/bin/env python3
import os
import argparse
import shutil
from dotenv import load_dotenv
from pdf_processor import PDFProcessor

def create_sample_env_if_missing():
    """Create .env file from .env.example if .env doesn't exist"""
    if not os.path.exists('.env') and os.path.exists('.env.example'):
        print("No .env file found. Creating from .env.example...")
        shutil.copy('.env.example', '.env')
        print("Created .env file. Please edit it with your API keys.")
        return True
    return False

def validate_api_keys(use_qwen):
    """Validate that API keys are set and not using placeholder values"""
    if use_qwen:
        key_name = "QWEN_API_KEY"
        api_key = os.getenv(key_name)
        placeholder = "your_qwen_api_key_here"
    else:
        key_name = "OPENAI_API_KEY"
        api_key = os.getenv(key_name)
        placeholder = "your_openai_api_key_here"
    
    if not api_key:
        print(f"Error: {key_name} environment variable is not set.")
        print(f"Please set it in your .env file or export it in your shell.")
        return False
    
    if api_key == placeholder:
        print(f"Error: {key_name} is still set to the placeholder value.")
        print(f"Please update it with your actual API key in the .env file.")
        return False
    
    return True

def create_sample_folder_if_missing(folder_path):
    """Create sample folder if it doesn't exist"""
    if not os.path.exists(folder_path):
        print(f"Creating sample folder: {folder_path}")
        os.makedirs(folder_path, exist_ok=True)
        print(f"Sample folder created. Please add PDF files to {folder_path}")
        return True
    return False

def main():
    """Main entry point for the example script"""
    parser = argparse.ArgumentParser(description="Example script for PDF processing")
    parser.add_argument("--folder", default="./pdf_files", help="Folder containing PDF files")
    parser.add_argument("--api-key", help="API key for OpenAI or Qwen")
    parser.add_argument("--use-qwen", action="store_true", help="Use Qwen API instead of OpenAI")
    parser.add_argument("--env-file", default=".env", help="Path to .env file")
    
    args = parser.parse_args()
    
    # Load environment variables from specified .env file
    if os.path.exists(args.env_file):
        load_dotenv(args.env_file)
        print(f"Loaded environment variables from {args.env_file}")
    else:
        # Try to create .env from .env.example if it doesn't exist
        if args.env_file == ".env":
            created = create_sample_env_if_missing()
            if created:
                load_dotenv()
            else:
                print(f"Warning: {args.env_file} not found and couldn't create from .env.example")
    
    # Create sample folder if it doesn't exist
    create_sample_folder_if_missing(args.folder)
    
    # Validate API keys
    if not args.api_key and not validate_api_keys(args.use_qwen):
        print("Please set the API key and try again.")
        return
    
    # Process PDFs
    print(f"Processing PDFs in folder: {args.folder}")
    print(f"Using {'Qwen' if args.use_qwen else 'OpenAI'} API")
    
    processor = PDFProcessor(
        folder_path=args.folder,
        api_key=args.api_key,
        use_qwen=args.use_qwen
    )
    
    results = processor.process_all()
    
    # Print summary
    print("\nProcessing Summary:")
    for doc_type, page_results in results.items():
        success_count = sum(1 for result in page_results if result.get("success", False))
        print(f"- {doc_type}: Processed {len(page_results)} pages, {success_count} successful")

if __name__ == "__main__":
    main()
