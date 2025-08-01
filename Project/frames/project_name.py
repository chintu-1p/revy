"""
Project name frame - Input project details
"""
import tkinter as tk
from tkinter import messagebox
from utils.validators import validate_project_name, validate_sample_count


class Project_name(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self._setup_ui()

    def _setup_ui(self):
        """Setup the user interface"""
        name_label = tk.Label(self, text="Enter Project Name:", font=("Arial", 12))
        name_label.pack(pady=(20, 5))

        self.project_name_var = tk.StringVar()
        name_entry = tk.Entry(self, textvariable=self.project_name_var, font=("Arial", 12), width=40)
        name_entry.pack()

        sample_label = tk.Label(self, text="Enter Number of Samples:", font=("Arial", 12))
        sample_label.pack(pady=(20, 5))

        self.sample_num_var = tk.IntVar()
        sample_num = tk.Entry(self, textvariable=self.sample_num_var, font=("Arial", 12), width=40)
        sample_num.pack()

        submit_btn = tk.Button(
            self, 
            text="Submit", 
            font=("Arial", 12),
            command=self.submit_project_info
        )
        submit_btn.pack(pady=10)

        back_btn = tk.Button(
            self, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack()

    def submit_project_info(self):
        """Submit and validate project information"""
        project_name = self.project_name_var.get().strip()
        sample_count = self.sample_num_var.get()
        
        # Validate inputs
        if not validate_project_name(project_name):
            messagebox.showwarning("Input Error", "Please enter a valid project name.")
            return
        
        if not validate_sample_count(sample_count):
            messagebox.showwarning("Input Error", "Please enter a valid number of samples (greater than 0).")
            return
        
        # Store project data
        self.controller.project_data['name'] = project_name
        self.controller.project_data['sample_count'] = sample_count
        
        # Initialize sample sheet names
        self.controller.project_data['sample_sheets'] = [f"Sample_{i+1}" for i in range(sample_count)]
        
        messagebox.showinfo("Project Created", f"Project '{project_name}' created with {sample_count} samples.")
        
        # Navigate to project type selection
        from frames.internal_project_type import InternalProjectType
        self.controller.show_frame(InternalProjectType)

    def get_project_name(self):
        """Get the current project name"""
        return self.project_name_var.get()

    def get_sample_count(self):
        """Get the current sample count"""
        return self.sample_num_var.get()
    
    def _go_back(self):
        """Navigate back to internal page"""
        from frames.internal_page import InternalPage
        self.controller.show_frame(InternalPage)