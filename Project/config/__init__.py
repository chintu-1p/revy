"""
Config package - Contains configuration settings and constants for the application.

This package includes:
- settings: Main application settings, paths, and configuration values
"""

# Import main settings
from .settings import (
    # Application settings
    APP_TITLE,
    APP_GEOMETRY,
    
    # File paths
    MASTER_TEMPLATE_PATH,
    BASE_PROJECT_DIR,
    
    # Excel settings
    EXCEL_COPY_RANGE,
    AUTO_FIT_COLUMN,
    
    # UI settings
    FONT_TITLE,
    FONT_HEADING,
    FONT_NORMAL,
    FONT_SMALL,
    COLOR_INFO,
    COLOR_SUCCESS,
    COLOR_ERROR,
    
    # Dimensions and spacing
    PADY_LARGE,
    PADY_MEDIUM,
    PADY_SMALL,
    BUTTON_WIDTH,
    ENTRY_WIDTH,
    
    # Validation settings
    MIN_SAMPLE_COUNT,
    MAX_SAMPLE_COUNT,
    MAX_PROJECT_NAME_LENGTH,
    
    # Template and utility functions
    PROJECT_DATA_TEMPLATE,
    get_setting
)

# Define what gets imported when using "from config import *"
__all__ = [
    # Application settings
    'APP_TITLE',
    'APP_GEOMETRY',
    
    # File paths
    'MASTER_TEMPLATE_PATH',
    'BASE_PROJECT_DIR',
    
    # Excel settings
    'EXCEL_COPY_RANGE',
    'AUTO_FIT_COLUMN',
    
    # UI settings
    'FONT_TITLE',
    'FONT_HEADING', 
    'FONT_NORMAL',
    'FONT_SMALL',
    'COLOR_INFO',
    'COLOR_SUCCESS',
    'COLOR_ERROR',
    
    # Dimensions
    'PADY_LARGE',
    'PADY_MEDIUM',
    'PADY_SMALL',
    'BUTTON_WIDTH',
    'ENTRY_WIDTH',
    
    # Validation
    'MIN_SAMPLE_COUNT',
    'MAX_SAMPLE_COUNT',
    'MAX_PROJECT_NAME_LENGTH',
    
    # Templates and functions
    'PROJECT_DATA_TEMPLATE',
    'get_setting'
]

# Package version
__version__ = '1.0.0'

# Package metadata
__author__ = 'Your Name'
__description__ = 'Configuration settings for project management application'

# Application info
APP_INFO = {
    'name': APP_TITLE,
    'version': __version__,
    'author': __author__,
    'description': __description__
}

def get_app_info():
    """
    Get application information
    
    Returns:
        dict: Application information dictionary
    """
    return APP_INFO.copy()

def validate_config():
    """
    Validate configuration settings
    
    Returns:
        bool: True if configuration is valid, False otherwise
    """
    import os
    
    # Check if master template exists
    if not os.path.exists(MASTER_TEMPLATE_PATH):
        print(f"Warning: Master template not found at {MASTER_TEMPLATE_PATH}")
        return False
    
    # Check if base directory exists or can be created
    try:
        os.makedirs(BASE_PROJECT_DIR, exist_ok=True)
    except Exception as e:
        print(f"Error: Cannot create base directory {BASE_PROJECT_DIR}: {e}")
        return False
    
    return True

# Validate configuration on import
_config_valid = validate_config()

def is_config_valid():
    """
    Check if configuration is valid
    
    Returns:
        bool: True if configuration is valid
    """
    return _config_valid