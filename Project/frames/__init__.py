"""
Frames package - Contains all GUI frame classes for the project management application.

This package includes:
- StartPage: Main entry point for project type selection
- InternalPage: Internal project options
- InternalProjectType: Project type selection after project creation
- Project_name: Project creation form
- Characterisation: Solid sample characterisation workflow
- BMPandCharacterisation: BMP and characterisation workflow
- SMA: SMA workflow
- Other: Other project types
"""

# Import all frame classes for easy access
from .start_page import StartPage
from .internal_page import InternalPage
from .internal_project_type import InternalProjectType
from .project_name import Project_name
from .characterisation import Characterisation
from .bmp_characterisation import BMPandCharacterisation
from .sma import SMA
from .other import Other

# Define what gets imported when using "from frames import *"
__all__ = [
    'StartPage',
    'InternalPage', 
    'InternalProjectType',
    'Project_name',
    'Characterisation',
    'BMPandCharacterisation',
    'SMA',
    'Other'
]

# Package version
__version__ = '1.0.0'

# Package metadata
__author__ = 'Your Name'
__description__ = 'GUI frames for project management application'