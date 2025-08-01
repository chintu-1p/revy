"""
BMP and Characterisation frame - OPTIMIZED VERSION
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
from utils.excel_handler import ExcelHandler
from utils.validators import validate_form_data
import requests


class BMPandCharacterisation(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.current_sample_index = 0
        self.excel_created = False
        self.sample_type = None
        self.excel_handler = ExcelHandler()
        self.form_widgets_created = False  # Track if form widgets are already created
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the user interface with scrolling capability"""
        # Create main canvas and scrollbar for scrolling
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Create a window inside the canvas, anchor north, and save its ID for resizing
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Bind canvas resize to update the width of the scrollable frame (center content)
        self.canvas.bind("<Configure>", self._resize_scrollable_frame)

        # Bind mousewheel to canvas for scrolling
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)

        # Make canvas focusable for keyboard scrolling
        self.canvas.bind("<Button-1>", lambda e: self.canvas.focus_set())
        self.canvas.bind("<Up>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Down>", lambda e: self.canvas.yview_scroll(1, "units"))
        self.canvas.bind("<Prior>", lambda e: self.canvas.yview_scroll(-1, "pages"))
        self.canvas.bind("<Next>", lambda e: self.canvas.yview_scroll(1, "pages"))

        # Now create all UI elements in the scrollable_frame instead of self
        label = tk.Label(self.scrollable_frame, text="BMP + Characterisation Page", font=("Arial", 16, "bold"))
        label.pack(pady=10)

        self.project_info_frame = tk.Frame(self.scrollable_frame)
        self.project_info_frame.pack(pady=5)

        # Sample type selection
        self.sample_type_var = tk.StringVar(value="Select")
        dropdown_label = tk.Label(self.scrollable_frame, text="Select Sample Type", font=("Arial", 12))
        dropdown_label.pack(pady=5)
        dropdown = ttk.OptionMenu(
            self.scrollable_frame,
            self.sample_type_var,
            "Select",
            "Solid",
            "Effluent",
            command=self.load_sample_selection
        )
        dropdown.pack()

        # Sample selection and management frames
        self.sample_selection_frame = tk.Frame(self.scrollable_frame)
        self.sample_selection_frame.pack(pady=10)

        self.sample_management_frame = tk.Frame(self.scrollable_frame)
        self.sample_management_frame.pack(pady=10)

        # Form frame for data entry
        self.form_frame = tk.Frame(self.scrollable_frame)
        self.form_frame.pack(pady=20)

        # Back button
        back_btn = tk.Button(
            self.scrollable_frame, 
            text="Back", 
            font=("Arial", 12),
            command=self._go_back
        )
        back_btn.pack(pady=10)

    def _resize_scrollable_frame(self, event):
        """Resize the scrollable_frame width to match canvas width to center content"""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        # Windows and MacOS
        if event.delta:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # Linux
        elif event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def _update_scroll_region(self):
        """Update the scroll region to encompass all widgets - OPTIMIZED"""
        # Use after_idle to batch scroll region updates
        if not hasattr(self, '_scroll_update_pending'):
            self._scroll_update_pending = True
            self.after_idle(self._do_scroll_update)

    def _do_scroll_update(self):
        """Actually perform the scroll region update"""
        try:
            self.scrollable_frame.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        finally:
            self._scroll_update_pending = False

    def tkraise(self):
        """Override tkraise to update project info when frame is shown"""
        super().tkraise()
        self.update_project_info()
        # Defer scroll region update to after widget creation
        self.after_idle(self._update_scroll_region)

    def update_project_info(self):
        """Update the project information display"""
        for widget in self.project_info_frame.winfo_children():
            widget.destroy()

        if self.controller.project_data['name']:
            info_text = (
                f"Project: {self.controller.project_data['name']} | "
                f"Samples: {self.controller.project_data['sample_count']}"
            )
            info_label = tk.Label(
                self.project_info_frame,
                text=info_text,
                font=("Arial", 10),
                fg="blue"
            )
            info_label.pack()

    def load_sample_selection(self, selected_value):
        """Load the appropriate interface based on sample type - OPTIMIZED"""
        self.sample_type = selected_value
        self._clear_frames()

        # Show loading message
        loading_label = tk.Label(
            self.sample_selection_frame, 
            text="Loading...", 
            font=("Arial", 12), 
            fg="blue"
        )
        loading_label.pack(pady=10)
        self.update()  # Force UI update

        try:
            if selected_value == "Solid":
                self._setup_solid_sample_interface()
            elif selected_value == "Effluent":
                self._setup_effluent_sample_interface()
        finally:
            # Remove loading message
            loading_label.destroy()
        
        # Defer scroll region update
        self.after_idle(self._update_scroll_region)

    def _clear_frames(self):
        """Clear all dynamic frames - OPTIMIZED"""
        frames_to_clear = [self.sample_selection_frame, self.sample_management_frame, self.form_frame]
        
        # Batch widget destruction for better performance
        for frame in frames_to_clear:
            for widget in frame.winfo_children():
                widget.destroy()
        
        # Reset form creation flag
        self.form_widgets_created = False

    def _setup_solid_sample_interface(self):
        """Setup interface for solid samples - OPTIMIZED"""
        if not self.excel_created:
            self._create_multi_sheet_excel("Solid")

        self._setup_sample_ui()
        # Defer form loading to improve perceived performance
        self.after(10, self.load_form_for_sample)

    def _setup_effluent_sample_interface(self):
        """Setup interface for effluent samples - OPTIMIZED"""
        if not self.excel_created:
            self._create_multi_sheet_excel("Effluent")

        self._setup_sample_ui()
        # Defer form loading to improve perceived performance
        self.after(10, self.load_form_for_sample)

    def _setup_sample_ui(self):
        """Setup the sample selection UI - OPTIMIZED"""
        sample_label = tk.Label(
            self.sample_selection_frame,
            text="Select Sample to Work On:",
            font=("Arial", 12, "bold")
        )
        sample_label.pack(pady=5)

        self.current_sample_var = tk.StringVar(
            value=self.controller.project_data['sample_sheets'][0]
        )
        sample_dropdown = ttk.OptionMenu(
            self.sample_selection_frame,
            self.current_sample_var,
            self.controller.project_data['sample_sheets'][0],
            *self.controller.project_data['sample_sheets'],
            command=self.change_sample
        )
        sample_dropdown.pack()

        # Create sample management buttons immediately
        self._setup_sample_management_buttons()

    def _setup_sample_management_buttons(self):
        """Setup sample management buttons"""
        btn_frame = tk.Frame(self.sample_management_frame)
        btn_frame.pack(pady=10)

        rename_btn = tk.Button(
            btn_frame,
            text="Rename Current Sample",
            font=("Arial", 10),
            command=self.rename_current_sample
        )
        rename_btn.pack(side=tk.LEFT, padx=5)

        progress_btn = tk.Button(
            btn_frame,
            text="Show Sample Progress",
            font=("Arial", 10),
            command=self.show_sample_progress
        )
        progress_btn.pack(side=tk.LEFT, padx=5)

    def _create_multi_sheet_excel(self, sample_type):
        """Create Excel workbook with multiple sheets for samples - OPTIMIZED"""
        try:
            project_data = self.controller.project_data
            
            # Show progress during Excel creation
            progress_label = tk.Label(
                self.sample_selection_frame, 
                text="Creating Excel file...", 
                font=("Arial", 10), 
                fg="orange"
            )
            progress_label.pack()
            self.update()  # Force UI update
            
            try:
                excel_path = self.excel_handler.create_bmp_workbook(
                    project_data['name'],
                    project_data['sample_count'],
                    project_data['sample_sheets'],
                    sample_type=sample_type
                )

                self.controller.project_data['excel_path'] = excel_path
                self.excel_created = True

                messagebox.showinfo(
                    "Excel Created",
                    f"BMP Excel file created with {project_data['sample_count']} sample sheets!"
                )
            finally:
                progress_label.destroy()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to create BMP Excel file:\n{e}")

    def change_sample(self, selected_sample):
        """Change the current sample being worked on - OPTIMIZED"""
        self.current_sample_index = self.controller.project_data['sample_sheets'].index(selected_sample)
        # Use after to prevent blocking the UI
        self.after(10, self.load_form_for_sample)

    def load_form_for_sample(self):
        """Load the form for the current sample - OPTIMIZED"""
        # Clear form frame efficiently
        for widget in self.form_frame.winfo_children():
            widget.destroy()

        current_sample = self.controller.project_data['sample_sheets'][self.current_sample_index]
        sample_info = tk.Label(
            self.form_frame,
            text=f"Working on: {current_sample}",
            font=("Arial", 12, "bold"),
            fg="green"
        )
        sample_info.pack(pady=10)

        # Create form entries based on sample type
        if self.sample_type == "Solid":
            self._create_form_entries_optimized("solid")
        elif self.sample_type == "Effluent":
            self._create_form_entries_optimized("effluent")

        # Create form buttons
        self._add_form_buttons()
        
        # Mark form as created
        self.form_widgets_created = True
        
        # Defer scroll region update
        self.after_idle(self._update_scroll_region)

    def _create_form_entries_optimized(self, sample_type):
        """Create form entries optimized for both solid and effluent samples"""
        self.entries = {}
        
        # Common labels for both types (they're the same anyway)
        labels = [
            ("Sample Type", "B1"),
            ("Sub Sample Type", "B2"),
            ("Sample Code", "D1"),
            ("Sample Receive Date", "D2"),
            ("Objective", "B3:D3"),
            ("Inoculum", "B4:D4"),
            ("Feed Sample", "B5:D5"),
            ("Trace", "B6:D6")
        ]
        
        # Batch create widgets for better performance
        self._generate_form_widgets_optimized(labels)

    def _generate_form_widgets_optimized(self, labels):
        """Generate form widgets based on labels - OPTIMIZED VERSION"""
        # Create all widgets in a batch to reduce individual pack() calls
        widgets_to_pack = []
        
        for text, cell in labels:
            # Create label
            label = tk.Label(self.form_frame, text=text, font=("Arial", 12))
            widgets_to_pack.append((label, {"anchor": "w", "padx": 20}))

            # Create entry widget
            if "Date" in text:
                entry = DateEntry(
                    self.form_frame, 
                    font=("Arial", 12), 
                    width=37, 
                    date_pattern='yyyy-mm-dd'
                )
            else:
                entry = tk.Entry(self.form_frame, font=("Arial", 12), width=40)

            widgets_to_pack.append((entry, {"padx": 20, "pady": 3}))
            self.entries[cell] = entry
        
        # Pack all widgets at once
        for widget, pack_options in widgets_to_pack:
            widget.pack(**pack_options)

    def _add_form_buttons(self):
        """Add form buttons for saving and navigation - OPTIMIZED"""
        # Create save button
        save_btn = tk.Button(
            self.form_frame,
            text="Save to Excel",
            font=("Arial", 12),
            command=self.save_current_sample
        )
        save_btn.pack(pady=10)

        # Create navigation frame and buttons
        nav_frame = tk.Frame(self.form_frame)
        nav_frame.pack(pady=5)

        # Only create buttons that are needed
        buttons_to_create = []
        
        if self.current_sample_index > 0:
            buttons_to_create.append(("← Previous Sample", self.go_to_previous_sample))

        if self.current_sample_index < len(self.controller.project_data['sample_sheets']) - 1:
            buttons_to_create.append(("Next Sample →", self.go_to_next_sample))

        # Create buttons efficiently
        for text, command in buttons_to_create:
            btn = tk.Button(
                nav_frame,
                text=text,
                font=("Arial", 10),
                command=command
            )
            btn.pack(side=tk.LEFT, padx=5)

    def go_to_previous_sample(self):
        """Navigate to the previous sample - OPTIMIZED"""
        if self.current_sample_index > 0:
            self.current_sample_index -= 1
            self.current_sample_var.set(
                self.controller.project_data['sample_sheets'][self.current_sample_index]
            )
            # Use after to prevent UI blocking
            self.after(10, self.load_form_for_sample)

    def go_to_next_sample(self):
        """Navigate to the next sample - OPTIMIZED"""
        if self.current_sample_index < len(self.controller.project_data['sample_sheets']) - 1:
            self.current_sample_index += 1
            self.current_sample_var.set(
                self.controller.project_data['sample_sheets'][self.current_sample_index]
            )
            # Use after to prevent UI blocking
            self.after(10, self.load_form_for_sample)

    def rename_current_sample(self):
        """Rename the current sample"""
        current_name = self.controller.project_data['sample_sheets'][self.current_sample_index]
        new_name = simpledialog.askstring("Rename Sample", f"Enter new name for '{current_name}':")

        if not new_name or not new_name.strip():
            return

        new_name = new_name.strip()

        if new_name in self.controller.project_data['sample_sheets']:
            messagebox.showwarning("Name Error", "A sample with this name already exists!")
            return

        try:
            self.excel_handler.rename_sheet(
                self.controller.project_data['excel_path'],
                current_name,
                new_name
            )

            self.controller.project_data['sample_sheets'][self.current_sample_index] = new_name
            self.current_sample_var.set(new_name)
            self.load_sample_selection(self.sample_type)

            messagebox.showinfo("Success", f"Sample renamed to '{new_name}'")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename sample:\n{e}")

    def show_sample_progress(self):
        """Show progress of all samples"""
        sample_list = "\n".join([f"• {name}" for name in self.controller.project_data['sample_sheets']])
        messagebox.showinfo("Sample Progress", f"All samples in project:\n\n{sample_list}")

    def save_current_sample(self):
        """Send current sample data to FastAPI for supervisor approval"""
        try:
            save_label = tk.Label(self.form_frame, text="Saving...", fg="orange", font=("Arial", 10))
            save_label.pack()
            self.update()

            try:
                form_data = {}
                for cell_range, entry in self.entries.items():
                    form_data[cell_range] = entry.get().strip()

                if not validate_form_data(form_data):
                    messagebox.showwarning("Input Error", "All fields must be filled.")
                    return

                current_sample = self.controller.project_data['sample_sheets'][self.current_sample_index]

                # Send to FastAPI backend
                try:
                    response = requests.post("http://127.0.0.1:8000/submit-sample", json={
                        "project_name": self.controller.project_data['name'],
                        "sample_name": current_sample,
                        "sample_type": self.sample_type,
                        "data": form_data
                    })

                    if response.status_code == 200:
                        messagebox.showinfo("Submitted", "Form sent for supervisor approval.")
                    else:
                        messagebox.showerror("Error", "Submission failed: " + response.text)

                except requests.exceptions.RequestException as e:
                    messagebox.showerror("Network Error", f"Could not reach backend:\n{e}")

            finally:
                save_label.destroy()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Unexpected error occurred:\n{e}")

    
    def _go_back(self):
        """Navigate back to internal project type page"""
        from frames.internal_project_type import InternalProjectType
        self.controller.show_frame(InternalProjectType)