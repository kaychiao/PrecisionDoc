from typing import Dict, Optional, Literal, Union, Any
from pydantic import BaseModel, Field, validator

class Margins(BaseModel):
    """Margin settings model"""
    left: float = Field(default=0.75, description="Left margin (inches)")
    right: float = Field(default=0.75, description="Right margin (inches)")
    top: float = Field(default=0.75, description="Top margin (inches)")
    bottom: float = Field(default=0.75, description="Bottom margin (inches)")

class PageSettings(BaseModel):
    """Word document page settings model"""
    orientation: Literal["portrait", "landscape"] = Field(
        default="landscape", 
        description="Page orientation, 'portrait' or 'landscape'"
    )
    margins: Optional[Margins] = Field(
        default=None, 
        description="Margin settings, dictionary with 'left', 'right', 'top', 'bottom' values in inches"
    )
    
    @validator('margins', pre=True)
    def validate_margins(cls, v):
        """Validate and convert margins parameter"""
        if v is None:
            return None
        if isinstance(v, dict):
            return Margins(**v)
        return v

class AISettings(BaseModel):
    """AI service settings model"""
    api_key: Optional[str] = Field(default=None, description="API key, if None will be loaded from environment variables")
    base_url: Optional[str] = Field(default=None, description="API base URL, if None will be loaded from environment variables")
    model: Optional[str] = Field(default=None, description="Model name to use, if None will be loaded from environment variables")
    
    class Config:
        """Pydantic configuration"""
        extra = "allow"  # Allow extra fields

def validate_ai_settings(ai_settings: Optional[Union[Dict[str, Any], AISettings]] = None, **legacy_params) -> AISettings:
    """
    Validate and convert AI settings parameters
    
    Args:
        ai_settings: AI settings dictionary or AISettings instance
        **legacy_params: Legacy parameters (api_key, base_url, model) for backward compatibility
        
    Returns:
        AISettings: Validated AI settings
    """
    # Handle legacy parameters
    if any(key in legacy_params for key in ['api_key', 'base_url', 'model']):
        if ai_settings is None:
            ai_settings = {}
        elif isinstance(ai_settings, AISettings):
            ai_settings = ai_settings.dict()
            
        # Add legacy parameters to ai_settings if they don't exist
        for key in ['api_key', 'base_url', 'model']:
            if key in legacy_params and (ai_settings.get(key) is None):
                ai_settings[key] = legacy_params[key]
    
    # If ai_settings is None, create an empty dictionary
    if ai_settings is None:
        ai_settings = {}
        
    # If ai_settings is already an AISettings instance, return it directly
    if isinstance(ai_settings, AISettings):
        return ai_settings
        
    # Otherwise, create a new AISettings instance
    return AISettings(**ai_settings)

def validate_page_settings(page_settings: Optional[Union[Dict[str, Any], PageSettings]] = None) -> Optional[PageSettings]:
    """
    Validate and convert page settings parameters
    
    Args:
        page_settings: Page settings dictionary or PageSettings instance
        
    Returns:
        Optional[PageSettings]: Validated page settings, or None if input is None
    """
    if page_settings is None:
        return None
        
    # If page_settings is already a PageSettings instance, return it directly
    if isinstance(page_settings, PageSettings):
        return page_settings
        
    # Otherwise, create a new PageSettings instance
    return PageSettings(**page_settings)
