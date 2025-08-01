"""
Internal page frame - Options for internal projects
"""
import tkinter as tk


class InternalPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        label = tk.Label(self, text="Internal Project Page", font=("Arial", 16))
        label.pack(pady=60)

        btn_create = tk.Button(
            self, 
            text="Create New Project", 
            font=("Arial", 14), 
            width=20,
            command=self._create_new_project
        )
        btn_create.pack(pady=10)

        back_btn = tk.Button(
            self, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack()


        btn_create = tk.Button(
            self, 
            text="Load Existing Project", 
            font=("Arial", 14), 
            width=20,
            command=self._load_project
        )
        btn_create.pack(pady=10)

        back_btn = tk.Button(
            self, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack()

        
    
    def _create_new_project(self):
        """Navigate to project name page"""
        from frames.project_name import Project_name
        self.controller.show_frame(Project_name)

    def _load_project(self):
        from frames.internal_project_type_load import InternalProjectTypeLoad
        self.controller.show_frame(InternalProjectTypeLoad)
    
    def _go_back(self):
        """Navigate back to start page"""
        from frames.start_page import StartPage
        self.controller.show_frame(StartPage)