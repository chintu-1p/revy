"""
Start page frame - Main entry point for project type selection
"""
import tkinter as tk


class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        label = tk.Label(self, text="Select Project Type", font=("Arial", 20, "bold"))
        label.pack(pady=50)

        btn_internal = tk.Button(
            self, 
            text="Internal Project", 
            font=("Arial", 14), 
            width=20,
            command=self._go_to_internal_page
        )
        btn_internal.pack(pady=10)
    
    def _go_to_internal_page(self):
        """Navigate to internal page"""
        from frames.internal_page import InternalPage
        self.controller.show_frame(InternalPage)