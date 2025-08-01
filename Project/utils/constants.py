"""
Application constants
"""

# Excel cell mappings for characterisation form
CHARACTERISATION_CELL_MAPPING = {
    "Sample Type": "B1",
    "Sub Sample Type": "B2", 
    "Sample Code": "D1",
    "Sample Receive Date": "D2",
    "Objective": "B3:D3",
    "Inoculum": "B4:D4",
    "Feed Sample": "B5:D5",
    "Trace": "B6:D6"
}

# Form field labels
CHARACTERISATION_FORM_LABELS = [
    ("Sample Type", "B1"),
    ("Sub Sample Type", "B2"),
    ("Sample Code", "D1"),
    ("Sample Receive Date", "D2"), 
    ("Objective", "B3:D3"),
    ("Inoculum", "B4:D4"),
    ("Feed Sample", "B5:D5"),
    ("Trace", "B6:D6")
]

# Sample types
SAMPLE_TYPES = ["Solid", "Effluent"]

# Project types
PROJECT_TYPES = ["Characterisation", "Characterisation and BMP", "SMA", "Other"]

# Maximum limits
MAX_PROJECT_NAME_LENGTH = 50
MAX_SAMPLE_COUNT = 100
MAX_SHEET_NAME_LENGTH = 31

# File extensions
EXCEL_EXTENSIONS = ['.xlsx', '.xls']

# Date format
DATE_FORMAT = '%Y-%m-%d'

# Error messages
ERROR_MESSAGES = {
    'invalid_project_name': 'Please enter a valid project name (no special characters, max 50 characters).',
    'invalid_sample_count': 'Please enter a valid number of samples (1-100).',
    'empty_fields': 'All fields must be filled.',
    'duplicate_sample_name': 'A sample with this name already exists!',
    'invalid_sample_name': 'Invalid sample name (no special characters, max 31 characters).',
    'excel_creation_failed': 'Failed to create Excel file.',
    'excel_save_failed': 'Failed to save to Excel file.',
    'sheet_rename_failed': 'Failed to rename sample sheet.'
}

# Success messages
SUCCESS_MESSAGES = {
    'project_created': 'Project created successfully!',
    'excel_created': 'Excel file created successfully!',
    'data_saved': 'Data saved successfully!',
    'sample_renamed': 'Sample renamed successfully!'
}