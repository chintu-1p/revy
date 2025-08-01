"""
Utils package - Contains utility functions and classes for the project management application.

This package includes:
- ExcelHandler: Excel file operations and management
- validators: Input validation functions
- constants: Application constants and configurations
"""

# Import main utility classes and functions
from .excel_handler import ExcelHandler
from .validators import (
    validate_project_name,
    validate_sample_count,
    validate_form_data,
    validate_date_format,
    validate_sample_name,
    validate_excel_path
)
from .constants import (
    CHARACTERISATION_CELL_MAPPING,
    CHARACTERISATION_FORM_LABELS,
    SAMPLE_TYPES,
    PROJECT_TYPES,
    ERROR_MESSAGES,
    SUCCESS_MESSAGES
)

# Define what gets imported when using "from utils import *"
__all__ = [
    # Classes
    'ExcelHandler',
    
    # Validation functions
    'validate_project_name',
    'validate_sample_count', 
    'validate_form_data',
    'validate_date_format',
    'validate_sample_name',
    'validate_excel_path',
    
    # Constants
    'CHARACTERISATION_CELL_MAPPING',
    'CHARACTERISATION_FORM_LABELS',
    'SAMPLE_TYPES',
    'PROJECT_TYPES',
    'ERROR_MESSAGES',
    'SUCCESS_MESSAGES'
]

# Package version
__version__ = '1.0.0'

# Package metadata
__author__ = 'Your Name'
__description__ = 'Utility functions and classes for project management'

# Initialize Excel handler instance for easy access
excel_handler = ExcelHandler()

def get_excel_handler():
    """
    Get a singleton instance of ExcelHandler
    
    Returns:
        ExcelHandler: Excel handler instance
    """
    return excel_handler