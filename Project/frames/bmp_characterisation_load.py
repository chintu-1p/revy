import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
from utils.excel_handler import ExcelHandler
from utils.validators import validate_form_data

class BMPandCharacterisationLoad(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.current_sample_index = 0
        self.excel_created = False
        self.sample_type = None
        self.excel_handler = ExcelHandler()
        self.form_widgets_created = False  # Track if form widgets are already created
            