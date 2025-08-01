"""
Enhanced Excel handling utilities for project management
"""
import xlwings as xw
import datetime
import os
import shutil
from typing import Dict, List, Any, Optional, Union
from config.settings import (
    MASTER_TEMPLATE_PATH, 
    BASE_PROJECT_DIR, 
    MASTER_TEMPLATE_PATH_EFFLUENT,
    MASTER_TEMPLATE_PATH_BMP_SOLID,
    MASTER_TEMPLATE_PATH_BMP_EFFLUENT
)


class ExcelHandler:
    """Enhanced Excel operations handler for the project management system"""

    def __init__(self):
        self.master_template_path = MASTER_TEMPLATE_PATH
        self.master_template_path_effluent = MASTER_TEMPLATE_PATH_EFFLUENT
        self.master_template_path_bmp_solid = MASTER_TEMPLATE_PATH_BMP_SOLID
        self.master_template_path_bmp_effluent = MASTER_TEMPLATE_PATH_BMP_EFFLUENT
        self.base_dir = BASE_PROJECT_DIR
        self._app_instance = None
        self._workbook_cache = {}

    def _get_app_instance(self, visible: bool = False) -> xw.App:
        """Get a reusable Excel application instance for better performance"""
        try:
            if self._app_instance is None:
                self._app_instance = xw.App(visible=visible)
                self._app_instance.display_alerts = False
                self._app_instance.screen_updating = False
            return self._app_instance
        except Exception:
            # Create new instance if current one is invalid
            self._app_instance = xw.App(visible=visible)
            self._app_instance.display_alerts = False
            self._app_instance.screen_updating = False
            return self._app_instance

    def _cleanup_app(self):
        """Clean up Excel application instance"""
        if self._app_instance:
            try:
                for wb in self._app_instance.books:
                    wb.close()
                self._app_instance.quit()
            except Exception:
                pass
            finally:
                self._app_instance = None
        self._workbook_cache.clear()

    def create_bmp_workbook(self, project_name: str, sample_count: int, 
                           sample_sheets: List[str], sample_type: str = "Solid") -> str:
        """
        Create a new Excel workbook with multiple sheets for BMP analysis

        Args:
            project_name: Name of the project
            sample_count: Number of samples
            sample_sheets: List of sheet names
            sample_type: "Solid" or "Effluent"

        Returns:
            Path to the created Excel file
        """
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{project_name}_{datetime.datetime.now().strftime('%Y%m%d')}"
        folder_path = os.path.join(self.base_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        new_file_path = os.path.join(folder_path, f"{project_name}_BMP_{sample_type}_{timestamp}.xlsx")

        app = self._get_app_instance()

        try:
            # Choose appropriate BMP template
            template_path = (
                self.master_template_path_bmp_solid if sample_type == "Solid" 
                else self.master_template_path_bmp_effluent
            )

            if not os.path.exists(template_path):
                raise FileNotFoundError(f"BMP template file not found: {template_path}")

            master_wb = app.books.open(template_path)
            master_sheet = master_wb.sheets[0]

            new_wb = app.books.add()
            
            # Create project summary sheet
            self._create_bmp_summary_sheet(new_wb, project_name, sample_count, sample_type)

            # Create sample sheets
            for i in range(sample_count):
                sheet_name = sample_sheets[i]

                if i == 0:
                    new_sheet = new_wb.sheets[0]
                    new_sheet.name = sheet_name
                else:
                    new_sheet = new_wb.sheets.add(name=sheet_name)

                # Copy template with better range detection
                self._copy_template_to_sheet(master_sheet, new_sheet)
                
                # Add BMP-specific formatting
                self._format_bmp_sheet(new_sheet, sheet_name)

            new_wb.save(new_file_path)
            master_wb.close()
            new_wb.close()

            # Create backup
            self._create_backup(new_file_path)

            return new_file_path

        except Exception as e:
            raise Exception(f"Failed to create BMP workbook: {e}")

    def create_characterisation_workbook(self, project_name: str, sample_count: int, 
                                       sample_sheets: List[str], sample_type: str = "Solid") -> str:
        """
        Create a new Excel workbook with multiple sheets for characterisation

        Args:
            project_name: Name of the project
            sample_count: Number of samples
            sample_sheets: List of sheet names
            sample_type: "Solid" or "Effluent"

        Returns:
            Path to the created Excel file
        """
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{project_name}_{datetime.datetime.now().strftime('%Y%m%d')}"
        folder_path = os.path.join(self.base_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        new_file_path = os.path.join(folder_path, f"{project_name}_characterisation_{sample_type}_{timestamp}.xlsx")

        app = self._get_app_instance()

        try:
            # Choose appropriate template
            template_path = (
                self.master_template_path if sample_type == "Solid"
                else self.master_template_path_effluent
            )

            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Characterisation template file not found: {template_path}")

            master_wb = app.books.open(template_path)
            master_sheet = master_wb.sheets[0]

            new_wb = app.books.add()
            
            # Create project summary sheet
            self._create_characterisation_summary_sheet(new_wb, project_name, sample_count, sample_type)

            # Create sample sheets
            for i in range(sample_count):
                sheet_name = sample_sheets[i]

                if i == 0:
                    new_sheet = new_wb.sheets[0]
                    new_sheet.name = sheet_name
                else:
                    new_sheet = new_wb.sheets.add(name=sheet_name)

                # Copy template with better range detection
                self._copy_template_to_sheet(master_sheet, new_sheet)

            new_wb.save(new_file_path)
            master_wb.close()
            new_wb.close()

            # Create backup
            self._create_backup(new_file_path)

            return new_file_path

        except Exception as e:
            raise Exception(f"Failed to create characterisation workbook: {e}")

    def _create_bmp_summary_sheet(self, workbook: xw.Book, project_name: str, 
                                 sample_count: int, sample_type: str):
        """Create a BMP project summary sheet"""
        try:
            summary_sheet = workbook.sheets.add("BMP_Summary", before=workbook.sheets[0])
            
            # Project information
            summary_sheet.range("A1").value = "BMP PROJECT SUMMARY"
            summary_sheet.range("A1").api.Font.Bold = True
            summary_sheet.range("A1").api.Font.Size = 16
            
            summary_data = [
                ["Project Name:", project_name],
                ["Project Type:", "BMP Analysis"],
                ["Sample Type:", sample_type],
                ["Total Samples:", sample_count],
                ["Created Date:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                ["Status:", "In Progress"]
            ]
            
            for i, (label, value) in enumerate(summary_data, start=3):
                summary_sheet.range(f"A{i}").value = label
                summary_sheet.range(f"B{i}").value = value
                summary_sheet.range(f"A{i}").api.Font.Bold = True
            
            # Auto-fit columns
            summary_sheet.range("A:B").api.EntireColumn.AutoFit()
            
        except Exception as e:
            print(f"Warning: Could not create BMP summary sheet: {e}")

    def _create_characterisation_summary_sheet(self, workbook: xw.Book, project_name: str, 
                                             sample_count: int, sample_type: str):
        """Create a characterisation project summary sheet"""
        try:
            summary_sheet = workbook.sheets.add("Char_Summary", before=workbook.sheets[0])
            
            # Project information
            summary_sheet.range("A1").value = "CHARACTERISATION PROJECT SUMMARY"
            summary_sheet.range("A1").api.Font.Bold = True
            summary_sheet.range("A1").api.Font.Size = 16
            
            summary_data = [
                ["Project Name:", project_name],
                ["Project Type:", "Characterisation"],
                ["Sample Type:", sample_type],
                ["Total Samples:", sample_count],
                ["Created Date:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                ["Status:", "In Progress"]
            ]
            
            for i, (label, value) in enumerate(summary_data, start=3):
                summary_sheet.range(f"A{i}").value = label
                summary_sheet.range(f"B{i}").value = value
                summary_sheet.range(f"A{i}").api.Font.Bold = True
            
            # Auto-fit columns
            summary_sheet.range("A:B").api.EntireColumn.AutoFit()
            
        except Exception as e:
            print(f"Warning: Could not create characterisation summary sheet: {e}")

    def _copy_template_to_sheet(self, master_sheet: xw.Sheet, target_sheet: xw.Sheet):
        """Copy template content to target sheet with smart range detection"""
        try:
            # Find the used range more intelligently
            used_range = master_sheet.used_range
            if used_range:
                # Copy the entire used range
                used_range.copy()
                target_sheet.range("A1").paste()
            else:
                # Fallback to fixed range
                master_sheet.range('A1:Z100').copy()
                target_sheet.range('A1').paste()
        except Exception:
            # Ultimate fallback
            master_sheet.range('A1:Z100').copy()
            target_sheet.range('A1').paste()

    def _format_bmp_sheet(self, sheet: xw.Sheet, sheet_name: str):
        """Apply BMP-specific formatting to a sheet"""
        try:
            # Add sheet identifier
            if sheet.range("A1").value is None:
                sheet.range("A1").value = f"BMP Analysis - {sheet_name}"
                sheet.range("A1").api.Font.Bold = True
                sheet.range("A1").api.Font.Size = 14
            
            # Auto-fit important columns
            for col in ["A", "B", "C", "D", "E"]:
                sheet.range(f"{col}:{col}").api.EntireColumn.AutoFit()
                
        except Exception as e:
            print(f"Warning: Could not format BMP sheet {sheet_name}: {e}")

    def _create_backup(self, file_path: str):
        """Create a backup of the Excel file"""
        try:
            backup_dir = os.path.join(os.path.dirname(file_path), "backups")
            os.makedirs(backup_dir, exist_ok=True)
            
            backup_name = f"backup_{os.path.basename(file_path)}"
            backup_path = os.path.join(backup_dir, backup_name)
            
            shutil.copy2(file_path, backup_path)
        except Exception as e:
            print(f"Warning: Could not create backup: {e}")

    def save_bmp_data(self, excel_path: str, sheet_name: str, data: Dict[str, Any]):
        """
        Save BMP data to a specific sheet with enhanced error handling

        Args:
            excel_path: Path to the Excel file
            sheet_name: Name of the sheet to save to
            data: Dictionary with cell ranges as keys and values to save
        """
        app = self._get_app_instance()

        try:
            wb = app.books.open(excel_path)
            sheet = wb.sheets[sheet_name]

            # Batch operations for better performance
            for cell_range, value in data.items():
                try:
                    if ':' in cell_range:
                        # Handle merged cells
                        sheet.range(cell_range).clear_contents()
                        try:
                            sheet.range(cell_range).api.UnMerge()
                        except:
                            pass  # Cell might not be merged
                        sheet.range(cell_range).value = value
                        sheet.range(cell_range).api.Merge()
                        sheet.range(cell_range).api.HorizontalAlignment = -4108  # Center alignment
                    else:
                        # Handle single cells
                        sheet.range(cell_range).clear_contents()
                        sheet.range(cell_range).value = value
                except Exception as e:
                    print(f"Warning: Could not save data to {cell_range}: {e}")

            # Auto-fit ALL columns after data entry
            self._auto_fit_all_columns(sheet)

            wb.save()
            wb.close()

        except Exception as e:
            raise Exception(f"Failed to save BMP data to Excel: {e}")

    def save_characterisation_data(self, excel_path: str, sheet_name: str, data: Dict[str, Any]):
        """
        Save characterisation data to a specific sheet with enhanced error handling

        Args:
            excel_path: Path to the Excel file
            sheet_name: Name of the sheet to save to
            data: Dictionary with cell ranges as keys and values to save
        """
        app = self._get_app_instance()

        try:
            wb = app.books.open(excel_path)
            sheet = wb.sheets[sheet_name]

            for cell_range, value in data.items():
                try:
                    if ':' in cell_range:
                        sheet.range(cell_range).clear_contents()
                        try:
                            sheet.range(cell_range).api.UnMerge()
                        except:
                            pass  # Cell might not be merged
                        sheet.range(cell_range).value = value
                        sheet.range(cell_range).api.Merge()
                        sheet.range(cell_range).api.HorizontalAlignment = -4108
                    else:
                        sheet.range(cell_range).clear_contents()
                        sheet.range(cell_range).value = value
                except Exception as e:
                    print(f"Warning: Could not save data to {cell_range}: {e}")

            # Auto-fit ALL columns after data entry
            self._auto_fit_all_columns(sheet)

            wb.save()
            wb.close()

        except Exception as e:
            raise Exception(f"Failed to save characterisation data to Excel: {e}")

    def _auto_fit_all_columns(self, sheet: xw.Sheet):
        """Auto-fit ALL columns with data for better readability"""
        try:
            # Method 1: Auto-fit entire worksheet columns (most comprehensive)
            sheet.api.Columns.AutoFit()
            
        except Exception as e:
            print(f"Warning: Could not auto-fit all columns with method 1: {e}")
            try:
                # Method 2: Auto-fit based on used range (fallback)
                used_range = sheet.used_range
                if used_range:
                    # Get the last column with data
                    last_col = used_range.last_cell.column
                    # Auto-fit columns from A to the last used column
                    for col_num in range(1, last_col + 1):
                        sheet.range((1, col_num)).api.EntireColumn.AutoFit()
                
            except Exception as e2:
                print(f"Warning: Could not auto-fit columns with method 2: {e2}")
                try:
                    # Method 3: Auto-fit common range (final fallback)
                    sheet.range("A:Z").api.EntireColumn.AutoFit()
                    
                except Exception as e3:
                    print(f"Warning: Could not auto-fit columns with method 3: {e3}")

    def _auto_fit_columns(self, sheet: xw.Sheet):
        """Legacy method - kept for backward compatibility, now calls the improved version"""
        self._auto_fit_all_columns(sheet)

    def _auto_fit_columns_smart(self, sheet: xw.Sheet, max_width: int = 50):
        """
        Smart auto-fit that limits column width to prevent extremely wide columns
        
        Args:
            sheet: The Excel sheet to auto-fit
            max_width: Maximum column width in characters (default: 50)
        """
        try:
            # First, auto-fit all columns
            sheet.api.Columns.AutoFit()
            
            # Then limit the width of columns that are too wide
            used_range = sheet.used_range
            if used_range:
                last_col = used_range.last_cell.column
                for col_num in range(1, last_col + 1):
                    col_range = sheet.range((1, col_num))
                    current_width = col_range.api.EntireColumn.ColumnWidth
                    if current_width > max_width:
                        col_range.api.EntireColumn.ColumnWidth = max_width
                        # Enable text wrapping for cells that might be cut off
                        col_range.api.EntireColumn.WrapText = True
                        
        except Exception as e:
            print(f"Warning: Could not perform smart auto-fit: {e}")
            # Fallback to basic auto-fit
            self._auto_fit_all_columns(sheet)

    def _auto_fit_specific_columns(self, sheet: xw.Sheet, column_list: List[str] = None):
        """
        Auto-fit specific columns only
        
        Args:
            sheet: The Excel sheet to auto-fit
            column_list: List of column letters to auto-fit (e.g., ['A', 'B', 'C'])
                        If None, auto-fits all columns with data
        """
        try:
            if column_list is None:
                # Auto-fit all columns with data
                self._auto_fit_all_columns(sheet)
            else:
                # Auto-fit only specified columns
                for col in column_list:
                    try:
                        sheet.range(f"{col}:{col}").api.EntireColumn.AutoFit()
                    except Exception as e:
                        print(f"Warning: Could not auto-fit column {col}: {e}")
                        
        except Exception as e:
            print(f"Warning: Could not auto-fit specific columns: {e}")

    def rename_sheet(self, excel_path: str, old_name: str, new_name: str):
        """
        Rename a sheet in the Excel workbook with validation

        Args:
            excel_path: Path to the Excel file
            old_name: Current name of the sheet
            new_name: New name for the sheet
        """
        app = self._get_app_instance()

        try:
            wb = app.books.open(excel_path)
            
            # Check if old sheet exists
            sheet_names = [sheet.name for sheet in wb.sheets]
            if old_name not in sheet_names:
                raise ValueError(f"Sheet '{old_name}' does not exist")
            
            # Check if new name already exists
            if new_name in sheet_names:
                raise ValueError(f"Sheet '{new_name}' already exists")
            
            sheet = wb.sheets[old_name]
            sheet.name = new_name
            wb.save()
            wb.close()

        except Exception as e:
            raise Exception(f"Failed to rename sheet: {e}")

    def get_sheet_data(self, excel_path: str, sheet_name: str, cell_ranges: List[str]) -> Dict[str, Any]:
        """
        Get data from specific cell ranges in a sheet with error handling

        Args:
            excel_path: Path to the Excel file
            sheet_name: Name of the sheet
            cell_ranges: List of cell ranges to read

        Returns:
            Dictionary with cell ranges as keys and their values
        """
        app = self._get_app_instance()

        try:
            wb = app.books.open(excel_path)
            sheet = wb.sheets[sheet_name]

            data = {}
            for cell_range in cell_ranges:
                try:
                    data[cell_range] = sheet.range(cell_range).value
                except Exception as e:
                    print(f"Warning: Could not read {cell_range}: {e}")
                    data[cell_range] = None

            wb.close()
            return data

        except Exception as e:
            raise Exception(f"Failed to read data from Excel: {e}")

    def check_sheet_exists(self, excel_path: str, sheet_name: str) -> bool:
        """
        Check if a sheet exists in the workbook

        Args:
            excel_path: Path to the Excel file
            sheet_name: Name of the sheet to check

        Returns:
            True if sheet exists, False otherwise
        """
        app = self._get_app_instance()

        try:
            wb = app.books.open(excel_path)
            sheet_names = [sheet.name for sheet in wb.sheets]
            wb.close()
            return sheet_name in sheet_names

        except Exception:
            return False

    def get_project_summary(self, excel_path: str) -> Dict[str, Any]:
        """
        Get project summary information from Excel file

        Args:
            excel_path: Path to the Excel file

        Returns:
            Dictionary with project summary data
        """
        app = self._get_app_instance()
        
        try:
            wb = app.books.open(excel_path)
            summary_data = {}
            
            # Look for summary sheets
            summary_sheets = [sheet for sheet in wb.sheets 
                            if 'summary' in sheet.name.lower()]
            
            if summary_sheets:
                summary_sheet = summary_sheets[0]
                # Extract summary information
                try:
                    summary_data['project_name'] = summary_sheet.range("B3").value
                    summary_data['project_type'] = summary_sheet.range("B4").value
                    summary_data['sample_type'] = summary_sheet.range("B5").value
                    summary_data['total_samples'] = summary_sheet.range("B6").value
                    summary_data['created_date'] = summary_sheet.range("B7").value
                    summary_data['status'] = summary_sheet.range("B8").value
                except Exception:
                    summary_data['error'] = "Could not read summary data"
            
            wb.close()
            return summary_data
            
        except Exception as e:
            return {'error': f"Failed to read project summary: {e}"}

    def __del__(self):
        """Cleanup when object is destroyed"""
        self._cleanup_app()