"""
Configuration settings for the project management application
"""
import os

# Application settings
APP_TITLE = "Project Management App"
APP_GEOMETRY = "800x650"

# File paths - UPDATE THESE PATHS FOR YOUR SYSTEM
MASTER_TEMPLATE_PATH = (
    r'C:\Users\Admin\Desktop\Master Template interface project\Master_template_solid_characterization.xlsx'
)

MASTER_TEMPLATE_PATH_EFFLUENT = ( 
    r'C:\Users\Admin\Desktop\Master Template interface project\Master_template_effluent_charaterization.xlsx'
)

# BMP Template Path - ADD YOUR BMP TEMPLATE PATH HERE
MASTER_TEMPLATE_PATH_BMP_SOLID = (
    r'C:\Users\Admin\Desktop\Master Template interface project\Master_template_BMP_Solid.xlsx'
)

MASTER_TEMPLATE_PATH_BMP_EFFLUENT = (
    r'C:\Users\Admin\Desktop\Master Template interface project\Master_template_BMP_effluent.xlsx'
)

BASE_PROJECT_DIR = (
    r'C:\Users\Admin\Desktop\Master Template interface project\Project\Internal Projects'
)

# Excel settings
EXCEL_COPY_RANGE = 'A1:Z100'  # Range to copy from master template
AUTO_FIT_COLUMN = 'D2'  # Column to auto-fit after saving

# UI settings
FONT_TITLE = ("Arial", 20, "bold")
FONT_HEADING = ("Arial", 16, "bold")
FONT_NORMAL = ("Arial", 12)
FONT_SMALL = ("Arial", 10)

# Colors
COLOR_INFO = "blue"
COLOR_SUCCESS = "green"
COLOR_ERROR = "red"

# Padding and spacing
PADY_LARGE = 60
PADY_MEDIUM = 20
PADY_SMALL = 10
PADY_TINY = 5

# Widget dimensions
BUTTON_WIDTH = 20
ENTRY_WIDTH = 40
DATE_ENTRY_WIDTH = 37

# Validation settings
MIN_SAMPLE_COUNT = 1
MAX_SAMPLE_COUNT = 100
MAX_PROJECT_NAME_LENGTH = 50
MAX_SHEET_NAME_LENGTH = 31

# File settings
ALLOWED_EXCEL_EXTENSIONS = ['.xlsx', '.xls']
DATE_FORMAT = '%Y-%m-%d'
TIMESTAMP_FORMAT = '%Y%m%d_%H%M%S'
FOLDER_DATE_FORMAT = '%Y%m%d'

# Excel application settings
EXCEL_VISIBLE = False
EXCEL_DISPLAY_ALERTS = False
EXCEL_SCREEN_UPDATING = False

# Project types
PROJECT_TYPES = {
    'BMP_SOLID': 'BMP - Solid Samples',
    'BMP_EFFLUENT': 'BMP - Effluent Samples',
    'CHARACTERISATION_SOLID': 'Characterisation - Solid Samples',
    'CHARACTERISATION_EFFLUENT': 'Characterisation - Effluent Samples',
    'BMP_CHARACTERISATION': 'BMP + Characterisation'
}

# Template mapping
TEMPLATE_MAPPING = {
    'BMP_SOLID': 'MASTER_TEMPLATE_PATH_BMP_SOLID',
    'BMP_EFFLUENT': 'MASTER_TEMPLATE_PATH_BMP_EFFLUENT',
    'CHARACTERISATION_SOLID': 'MASTER_TEMPLATE_PATH',
    'CHARACTERISATION_EFFLUENT': 'MASTER_TEMPLATE_PATH_EFFLUENT'
}

# Project data structure template
PROJECT_DATA_TEMPLATE = {
    'name': '',
    'sample_count': 0,
    'excel_path': '',
    'sample_sheets': [],
    'project_type': '',
    'sample_type': ''
}

# Environment settings
def get_setting(key, default=None):
    """Get setting from environment variables or use default"""
    return os.environ.get(key, default)

def get_template_path(project_type, sample_type):
    """
    Get the appropriate template path based on project and sample type
    
    Args:
        project_type (str): 'BMP' or 'Characterisation'
        sample_type (str): 'Solid' or 'Effluent'
    
    Returns:
        str: Path to the appropriate template file
    """
    template_key = f"{project_type.upper()}_{sample_type.upper()}"
    
    if template_key == 'BMP_SOLID':
        return MASTER_TEMPLATE_PATH_BMP_SOLID
    elif template_key == 'BMP_EFFLUENT':
        return MASTER_TEMPLATE_PATH_BMP_EFFLUENT
    elif template_key == 'CHARACTERISATION_SOLID':
        return MASTER_TEMPLATE_PATH
    elif template_key == 'CHARACTERISATION_EFFLUENT':
        return MASTER_TEMPLATE_PATH_EFFLUENT
    else:
        # Default fallback
        return MASTER_TEMPLATE_PATH

def validate_template_paths():
    """
    Validate that all template paths exist
    
    Returns:
        dict: Dictionary with validation results
    """
    paths = {
        'Characterisation Solid': MASTER_TEMPLATE_PATH,
        'Characterisation Effluent': MASTER_TEMPLATE_PATH_EFFLUENT,
        'BMP Solid': MASTER_TEMPLATE_PATH_BMP_SOLID,
        'BMP Effluent': MASTER_TEMPLATE_PATH_BMP_EFFLUENT
    }
    
    results = {}
    for name, path in paths.items():
        results[name] = {
            'path': path,
            'exists': os.path.exists(path),
            'readable': os.path.exists(path) and os.access(path, os.R_OK)
        }
    
    return results

def get_project_folder_path(project_name):
    """
    Get the full path for a project folder
    
    Args:
        project_name (str): Name of the project
    
    Returns:
        str: Full path to the project folder
    """
    from datetime import datetime
    folder_name = f"{project_name}_{datetime.now().strftime(FOLDER_DATE_FORMAT)}"
    return os.path.join(BASE_PROJECT_DIR, folder_name)

# Override paths from environment if available
MASTER_TEMPLATE_PATH = get_setting('MASTER_TEMPLATE_PATH', MASTER_TEMPLATE_PATH)
MASTER_TEMPLATE_PATH_EFFLUENT = get_setting('MASTER_TEMPLATE_PATH_EFFLUENT', MASTER_TEMPLATE_PATH_EFFLUENT)
MASTER_TEMPLATE_PATH_BMP_SOLID = get_setting('MASTER_TEMPLATE_PATH_BMP_SOLID', MASTER_TEMPLATE_PATH_BMP_SOLID)
MASTER_TEMPLATE_PATH_BMP_EFFLUENT = get_setting('MASTER_TEMPLATE_PATH_BMP_EFFLUENT', MASTER_TEMPLATE_PATH_BMP_EFFLUENT)
BASE_PROJECT_DIR = get_setting('BASE_PROJECT_DIR', BASE_PROJECT_DIR)