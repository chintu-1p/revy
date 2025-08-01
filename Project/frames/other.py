"""
Other project frame
"""
import tkinter as tk


class Other(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the user interface"""
        label = tk.Label(self, text="Other Project Page", font=("Arial", 16))
        label.pack(pady=60)

        back_btn = tk.Button(
            self, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack()
    
    def _go_back(self):
        """Navigate back to internal project type page"""
        from frames.internal_project_type import InternalProjectType
        self.controller.show_frame(InternalProjectType)