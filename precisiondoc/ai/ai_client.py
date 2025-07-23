import os
import openai
import base64
from typing import Dict, Optional

# Import Qwen client
from .qwen_api import QwenClient
from ..utils.log_utils import setup_logger

# Setup logger for this module
logger = setup_logger(__name__)

class AIClient:
    """Client for interacting with AI APIs (OpenAI or Qwen)"""
    
    def __init__(self, api_key: str = None, use_qwen: bool = False):
        """
        Initialize the AI client.
        
        Args:
            api_key: API key for OpenAI or Qwen. If None, will try to load from environment variables.
            use_qwen: If True, use Qwen API instead of OpenAI
        """
        self.use_qwen = use_qwen
        
        # If API key is not provided, try to load from environment variables
        if api_key is None:
            if use_qwen:
                self.api_key = os.getenv("QWEN_API_KEY")
                if not self.api_key:
                    raise ValueError("QWEN_API_KEY environment variable not set")
            else:
                self.api_key = os.getenv("OPENAI_API_KEY")
                if not self.api_key:
                    raise ValueError("OPENAI_API_KEY environment variable not set")
        else:
            self.api_key = api_key
            
        # Initialize Qwen client if needed
        if self.use_qwen:
            self.qwen_client = QwenClient(api_key=self.api_key)
    
    def identify_page_type(self, text: str, image_path: Optional[str] = None) -> Dict:
        """
        Identify the type of page (content, table of contents, references, etc.)
        
        Args:
            text: Text extracted from the page
            image_path: Optional path to the page image
            
        Returns:
            Dictionary with page type information
        """
        prompt = """
        请判断这个PDF页面的类型。可能的类型包括：
        1. 目录页 (table_of_contents) - 包含章节列表和页码
        2. 参考文献页 (references) - 包含引用的文献列表
        3. 内容页 (content) - 包含实际的医疗指南内容
        
        请仅返回一个单词作为页面类型：table_of_contents、references 或 content
        """
        
        if self.use_qwen and image_path:
            # Use image-based identification with Qwen
            response = self._call_qwen_api_with_image(prompt, image_path)
        else:
            # Use text-based identification
            prompt += f"\n\n页面文本内容：\n{text[:1000]}..."  # Limit text length
            if self.use_qwen:
                response = self._call_qwen_api(prompt)
            else:
                response = self._call_openai_api(prompt)
        
        # Extract page type from response
        if response.get("success", False):
            content = response.get("content", "").lower().strip()
            if "table_of_contents" in content or "目录" in content:
                page_type = "table_of_contents"
            elif "references" in content or "参考文献" in content:
                page_type = "references"
            else:
                page_type = "content"
            
            return {"success": True, "page_type": page_type}
        else:
            # Default to content if identification fails
            return {"success": False, "page_type": "content", "error": response.get("error")}
    
    def process_text(self, text: str) -> Dict:
        """
        Process text with AI (OpenAI or Qwen).
        
        Args:
            text: Text to process
            
        Returns:
            Dictionary containing the AI's response
        """
        # Prepare prompt for AI - specialized for precision medicine
        prompt = f"""
        请分析以下医疗文本，判断文字中是否能提供精准医疗相关的用药证据，即是否涉及某个基因或基因变异与特定肿瘤疾病在使用某种药物（或药物组合）后的疗效（敏感性/耐药性等）或疗效预测关系。

        文本内容：
        {text}

        如果是，请提取并输出如下结构化证据信息（未提及的字段请填 null）：
        - 相关基因（symbol）及变异（alteration）
        - 疾病的中文名和英文名
        - 药物中文名和英文名，及药物组合（如果有）
        - 证据等级（A/B/C/D）、响应性（敏感/耐药）、证据类型
            A1(FDA-approved therapies)
            A2(Professional guidelines)
            B(Well-powered studies with consensus)
            C1(Multiple small studies with some consensus)
            C2(inclusion criteria for CT)
            C3(A-evidence for a different Ca)
            D1(Cases)
            D2(Preclinical)

        输出格式为 JSON，包含以下字段：
        {
          "text": "原文提取的文本",
          "is_precision_evidence": true/false,
          "symbol": "基因符号",
          "alteration": "基因变异",
          "disease_name_cn": "疾病中文名",
          "disease_name_en": "疾病英文名",
          "drug_name_cn": "药物中文名",
          "drug_name_en": "药物英文名",
          "drug_combination": "药物组合",
          "evidence_level": "证据等级",
          "response_type": "敏感/耐药",
          "evidence_type": "证据类型"
        }

        如果该文本内容不涉及基因变异与疾病的药物疗效关系，请只返回：{"is_precision_evidence": false}
        """
        
        if self.use_qwen:
            return self._call_qwen_api(prompt)
        else:
            return self._call_openai_api(prompt)
    
    def process_image(self, image_path: str) -> Dict:
        """
        Process an image of a PDF page with AI.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Dictionary containing the AI's response
        """
        if not self.use_qwen:
            logger.warning("Image processing is only available with Qwen API. Falling back to text extraction.")
            return {"success": False, "error": "Image processing requires Qwen API"}
        
        # Prepare prompt for AI - specialized for precision medicine
        prompt = """
        请先把该图片中的文字原文提取出来，判断文字中是否能提供精准医疗相关的用药证据，即是否涉及某个基因或基因变异与特定肿瘤疾病在使用某种药物（或药物组合）后的疗效（敏感性/耐药性等）或疗效预测关系。

        如果是，请提取并输出如下结构化证据信息（未提及的字段请填 null）：
        - 相关基因（symbol）及变异（alteration）
        - 疾病的中文名和英文名
        - 药物中文名和英文名，及药物组合（如果有）
        - 证据等级（A/B/C/D）、响应性（敏感/耐药）、证据类型
            A1(FDA-approved therapies)
            A2(Professional guidelines)
            B(Well-powered studies with consensus)
            C1(Multiple small studies with some consensus)
            C2(inclusion criteria for CT)
            C3(A-evidence for a different Ca)
            D1(Cases)
            D2(Preclinical)

        输出格式为 JSON，包含以下字段：
        {
          "text": "原文提取的文本",
          "is_precision_evidence": true/false,
          "symbol": "基因符号",
          "alteration": "基因变异",
          "disease_name_cn": "疾病中文名",
          "disease_name_en": "疾病英文名",
          "drug_name_cn": "药物中文名",
          "drug_name_en": "药物英文名",
          "drug_combination": "药物组合",
          "evidence_level": "证据等级",
          "response_type": "敏感/耐药",
          "evidence_type": "证据类型"
        }

        如果该图片内容不涉及基因变异与疾病的药物疗效关系，请只返回：{"is_precision_evidence": false}
        """
        
        return self._call_qwen_api_with_image(prompt, image_path)
    
    def _call_openai_api(self, prompt: str) -> Dict:
        """
        Call OpenAI API with the given prompt.
        
        Args:
            prompt: Prompt to send to OpenAI
            
        Returns:
            Dictionary containing the response
        """
        openai.api_key = self.api_key
        
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a medical document analysis assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            return {"success": True, "content": response.choices[0].message.content}
        except Exception as e:
            logger.error(f"Error calling OpenAI API: {str(e)}")
            return {"success": False, "error": str(e)}
    
    def _call_qwen_api(self, prompt: str) -> Dict:
        """
        Call Qwen API with the given prompt.
        
        Args:
            prompt: Prompt to send to Qwen
            
        Returns:
            Dictionary containing the response
        """
        try:
            response = self.qwen_client.chat(
                prompt=prompt,
                system_prompt="You are a medical document analysis assistant."
            )
            return {"success": True, "content": response}
        except Exception as e:
            logger.error(f"Error calling Qwen API: {str(e)}")
            return {"success": False, "error": str(e)}
    
    def _call_qwen_api_with_image(self, prompt: str, image_path: str) -> Dict:
        """
        Call Qwen API with the given prompt and image.
        
        Args:
            prompt: Prompt to send to Qwen
            image_path: Path to the image file
            
        Returns:
            Dictionary containing the response
        """
        try:
            # Read image file and encode as base64
            with open(image_path, "rb") as image_file:
                image_data = base64.b64encode(image_file.read()).decode("utf-8")
            
            # Call Qwen API with image
            response = self.qwen_client.chat_with_image(
                prompt=prompt,
                image_data=image_data,
                system_prompt="You are a medical document analysis assistant specialized in analyzing medical documents and images."
            )
            return {"success": True, "content": response}
        except Exception as e:
            logger.error(f"Error calling Qwen API with image: {str(e)}")
            return {"success": False, "error": str(e)}
