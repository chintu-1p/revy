"""
Internal project type frame - Select project type after creating project
"""
import tkinter as tk


class InternalProjectTypeLoad(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        self._setup_ui()

    def _setup_ui(self):
        """Setup the user interface"""
        label = tk.Label(self, text="Select Project Type", font=("Arial", 16))
        label.pack(pady=40)

        # Show project info
        self.info_frame = tk.Frame(self)
        self.info_frame.pack(pady=10)

        btn_char = tk.Button(
            self, 
            text="Characterisation", 
            font=("Arial", 14), 
            width=30,
            command=self._go_to_characterisation
        )
        btn_char.pack(pady=10)

        btn_char_bmp = tk.Button(
            self, 
            text="Characterisation and BMP", 
            font=("Arial", 14), 
            width=30,
            command=self._go_to_bmp_characterisation
        )
        btn_char_bmp.pack(pady=10)

        btn_sma = tk.Button(
            self, 
            text="SMA", 
            font=("Arial", 14), 
            width=30,
            command=self._go_to_sma
        )
        btn_sma.pack(pady=10)

        btn_other = tk.Button(
            self, 
            text="Other", 
            font=("Arial", 14), 
            width=30,
            command=self._go_to_other
        )
        btn_other.pack(pady=10)

        back_btn = tk.Button(
            self, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack(pady=20)

    def tkraise(self):
        """Override to update project info when frame is raised"""
        super().tkraise()
        self.update_project_info()

    def update_project_info(self):
        """Update project information display"""
        # Clear previous info
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        # Show current project info
        if self.controller.project_data['name']:
            info_text = (f"Project: {self.controller.project_data['name']} | "
                        f"Samples: {self.controller.project_data['sample_count']}")
            info_label = tk.Label(self.info_frame, text=info_text, font=("Arial", 10), fg="blue")
            info_label.pack()
    
    def _go_to_characterisation(self):
        """Navigate to characterisation page"""
        from frames.characterisation_load import Characterisation
        self.controller.show_frame(Characterisation)
    
    def _go_to_bmp_characterisation(self):
        """Navigate to BMP characterisation page"""
        from frames.bmp_characterisation_load import BMPandCharacterisation
        self.controller.show_frame(BMPandCharacterisation)
    
    def _go_to_sma(self):
        """Navigate to SMA page"""
        from frames.sma_load import SMA
        self.controller.show_frame(SMA)
    
    def _go_to_other(self):
        """Navigate to other page"""
        from frames.other_load import Other
        self.controller.show_frame(Other)
    
    def _go_back(self):
        """Navigate back to internal page"""
        from frames.internal_page import InternalPage
        self.controller.show_frame(InternalPage)