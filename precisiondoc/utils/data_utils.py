import os
import json
import pandas as pd
from typing import Dict, List, Any
from datetime import datetime

from .log_utils import setup_logger

# Setup logger for this module
logger = setup_logger(__name__)

class DataUtils:
    """Data processing utility class"""
    
    @staticmethod
    def handle_json_file(results: Dict[str, List[Dict]], output_file: str = None, output_folder: str = None) -> str:
        """
        Handle JSON file read/write operations
        
        Args:
            results: Result data
            output_file: Output file path
            output_folder: Output folder path
            
        Returns:
            Processed output file path
        """
        if not output_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_folder, f"all_results_{timestamp}.json")
            
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
        # For testing purposes, load from existing file
        json_file = os.path.join(output_folder, "all_results_20250717_094112.json")
        if os.path.exists(json_file):
            with open(json_file, 'r', encoding='utf-8') as f:
                results = json.load(f)
            output_file = json_file
            logger.info(f"Loaded existing results from {json_file}")
        else:
            # Save results to JSON file
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            logger.info(f"Consolidated results saved to {output_file}")
        
        return output_file
    
    @staticmethod
    def convert_to_flat_structure(results: Dict[str, List[Dict]]) -> List[Dict]:
        """
        Convert nested JSON structure to flat data structure
        
        Args:
            results: Nested result data
            
        Returns:
            List of flattened data rows
        """
        all_rows = []
        
        # Process each document type
        for doc_type, pages in results.items():
            for page_idx, page_data in enumerate(pages):
                # Create new row with document type and page number
                row = {'document_type': doc_type, 'page_number': page_idx + 1}
                
                # Add basic page data (all top-level keys except 'content')
                if isinstance(page_data, dict):
                    for key, value in page_data.items():
                        if key != 'content':
                            row[key] = value
                    
                    # If content field exists, process it
                    if 'content' in page_data:
                        content = page_data['content']
                        
                        # Process JSON and Markdown content
                        DataUtils.process_json_content(content, row)
                        DataUtils.process_markdown_sections(content, row)
                
                all_rows.append(row)
        
        return all_rows
    
    @staticmethod
    def process_json_content(content: str, row: Dict) -> None:
        """
        Extract JSON data from markdown code blocks
        
        Args:
            content: Content text
            row: Data row to update
        """
        if not isinstance(content, str) or '```json' not in content:
            return
            
        try:
            # Extract JSON strings
            json_parts = content.split('```json')
            for part_idx, part in enumerate(json_parts):
                if part_idx > 0:  # Skip first part (content before the first ```json)
                    json_str = part.split('```')[0].strip()
                    try:
                        json_data = json.loads(json_str)
                        # Add JSON data to row
                        if isinstance(json_data, dict):
                            for key, value in json_data.items():
                                row[key] = value
                    except json.JSONDecodeError:
                        logger.warning(f"Could not parse JSON part {part_idx} from content")
        except Exception as e:
            logger.warning(f"Error processing JSON in content: {str(e)}")
    
    @staticmethod
    def process_markdown_sections(content: str, row: Dict) -> None:
        """
        Process markdown format content and headings
        
        Args:
            content: Content text
            row: Data row to update
        """
        if not isinstance(content, str):
            return
            
        # Skip parts already processed as JSON
        non_json_content = content
        if '```json' in content:
            # Remove JSON code blocks
            parts = content.split('```json')
            non_json_content = parts[0]  # Text before first JSON block
            for i in range(1, len(parts)):
                if '```' in parts[i]:
                    non_json_content += parts[i].split('```', 1)[1]
        
        # Process remaining content by sections
        lines = non_json_content.split('\n')
        current_heading = None
        current_content = []
        
        for line in lines:
            if line.strip().startswith('###'):
                # If exists, save previous section
                if current_heading:
                    section_content = '\n'.join(current_content).strip()
                    if section_content:
                        row[current_heading] = section_content
                
                # Start new section
                current_heading = line.strip().replace('###', '').strip()
                current_content = []
            else:
                if current_heading:
                    current_content.append(line)
        
        # Save last section
        if current_heading and current_content:
            section_content = '\n'.join(current_content).strip()
            if section_content:
                row[current_heading] = section_content
    
    @staticmethod
    def save_to_excel(all_rows: List[Dict], excel_file: str) -> None:
        """
        Create DataFrame and save to Excel
        
        Args:
            all_rows: Data rows to save
            excel_file: Excel file path
        """
        if all_rows:
            df = pd.DataFrame(all_rows)
            
            # Save to Excel
            df.to_excel(excel_file, index=False)
            logger.info(f"Consolidated results saved to Excel file: {excel_file}")
        else:
            logger.warning("No data to save to Excel file")
