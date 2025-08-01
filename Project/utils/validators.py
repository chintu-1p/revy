"""
Input validation utilities
"""
import re
from datetime import datetime


def validate_project_name(project_name):
    """
    Validate project name
    
    Args:
        project_name (str): Project name to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not project_name or not project_name.strip():
        return False
    
    # Check for invalid characters in file names
    invalid_chars = r'[<>:"/\\|?*]'
    if re.search(invalid_chars, project_name):
        return False
    
    # Check length
    if len(project_name.strip()) > 50:
        return False
        
    return True


def validate_sample_count(sample_count):
    """
    Validate sample count
    
    Args:
        sample_count (int): Number of samples
        
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        count = int(sample_count)
        return count > 0 and count <= 100  # Reasonable upper limit
    except (ValueError, TypeError):
        return False


def validate_form_data(form_data):
    """
    Validate form data for completeness
    
    Args:
        form_data (dict): Dictionary of form field values
        
    Returns:
        bool: True if all fields are filled, False otherwise
    """
    for key, value in form_data.items():
        if not value or not str(value).strip():
            return False
    return True


def validate_date_format(date_string, date_format='%Y-%m-%d'):
    """
    Validate date format
    
    Args:
        date_string (str): Date string to validate
        date_format (str): Expected date format
        
    Returns:
        bool: True if valid date format, False otherwise
    """
    try:
        datetime.strptime(date_string, date_format)
        return True
    except ValueError:
        return False


def validate_sample_name(sample_name, existing_names):
    """
    Validate sample name for uniqueness and format
    
    Args:
        sample_name (str): Sample name to validate
        existing_names (list): List of existing sample names
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not sample_name or not sample_name.strip():
        return False
    
    # Check for invalid characters in sheet names
    invalid_chars = r'[\\/*?:\[\]]'
    if re.search(invalid_chars, sample_name):
        return False
    
    # Check length (Excel sheet name limit)
    if len(sample_name.strip()) > 31:
        return False
    
    # Check uniqueness
    if sample_name.strip() in existing_names:
        return False
        
    return True


def validate_excel_path(file_path):
    """
    Validate Excel file path
    
    Args:
        file_path (str): Path to Excel file
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not file_path:
        return False
    
    # Check file extension
    valid_extensions = ['.xlsx', '.xls']
    if not any(file_path.lower().endswith(ext) for ext in valid_extensions):
        return False
    
    return True