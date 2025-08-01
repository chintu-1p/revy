"""
Main application entry point
"""
import tkinter as tk
from frames.start_page import StartPage
from frames.internal_page import InternalPage
from frames.internal_project_type import InternalProjectType
from frames.internal_project_type_load import InternalProjectTypeLoad
from frames.project_name import Project_name
from frames.characterisation import Characterisation
from frames.bmp_characterisation import BMPandCharacterisation
from frames.bmp_characterisation_load import BMPandCharacterisationLoad
from frames.sma import SMA
from frames.other import Other



class ProjectManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Project Management App")
        self.geometry("800x650")
        self.resizable(True, True)

        # Store project-wide data
        self.project_data = {
            'name': '',
            'sample_count': 0,
            'excel_path': '',
            'sample_sheets': []
        }

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (StartPage, InternalPage, InternalProjectType,InternalProjectTypeLoad,BMPandCharacterisationLoad, Project_name,
                  Characterisation, BMPandCharacterisation, SMA, Other):
            frame = F(parent=self.container, controller=self)
            self.frames[F] = frame
            frame.place(relwidth=1, relheight=1)

        self.show_frame(StartPage)

    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()


if __name__ == "__main__":
    app = ProjectManagementApp()
    app.mainloop()