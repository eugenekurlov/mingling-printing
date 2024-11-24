import os
import sys
import io
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pymupdf
from PIL import Image, ImageTk

from create_file import add_images_to_pdf_in_grid, create_pdf_with_best_orientation_images, extract_and_merge_pdfs
from printer_utils import PrinterManager


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class App(tk.Tk):
    def __init__(self):
        # main setup
        super().__init__()
        self.title("Mingling Printing")
        self.iconbitmap(resource_path("printer.ico"))
        
        # Set the desired window size
        window_width = 400
        window_height = 300

        # Get screen dimensions
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate the center position
        center_x = (screen_width - window_width) // 2
        center_y = (screen_height - window_height) // 2

        # Position the window at the center
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        # widgets
        self.start_menu = StartMenu(self)

        # run
        self.mainloop()


class StartMenu(tk.Frame):
    def __init__(self, master_parent):
        super().__init__(master_parent)

        # Buttons to open Images and PDF windows
        tk.Button(self, text="Images", command=self.open_images_window).pack(expand=True, fill="both")
        tk.Button(self, text="PDFs", command=self.open_pdf_window).pack(expand=True, fill="both")
        self.pack(expand=True, fill="both", padx=20, pady=20)

    def open_images_window(self):
        # Open a new ImagesWindow
        ImagesWindow(self.master)

    def open_pdf_window(self):
        # Open a new PDFWindow
        PDFWindow(self.master)


class ImagesWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Files Configuration Window")
        self.iconbitmap(resource_path("printer.ico"))
        self.geometry("+50+50")

        # Sample widgets for the Images Window

        self.output_path = tk.BooleanVar()
        self.multiple_pages = tk.BooleanVar()
        self.angles_needed = tk.BooleanVar()
        self.orientation = tk.StringVar(value="portrait")
        self.best_orientation = tk.BooleanVar()
        self.columns_number = tk.IntVar(value=1)
        self.rows_number = tk.IntVar(value=1)
        self.page_margin = tk.IntVar(value=0)
        self.image_margin = tk.IntVar(value=0)
        self.output_path_file = None
        self.image_paths = []
        self.thumbnails = []  # Store thumbnails to avoid garbage collection
        self.angle_entries = {}

        self.angles_needed.trace_add("write", self.update_angles_needed)
        self.best_orientation.trace_add("write", self.update_best_orientation)

        self.input_interface = InputImagesInterface(self)
        self.input_interface.grid(row=0, column=0, sticky="nsew")

        self.output_options = OutputOptionsInterface(self)
        self.output_options.grid(row=2, column=0, sticky="w")

        self.multi_page = MultiplePagesInterface(self)
        self.multi_page.grid(row=3, column=0, pady=10, sticky="w")

        self.page_orientation = PagesOrientationInterface(self)
        self.page_orientation.grid(row=4, column=0, pady=10, sticky="w")

        self.margins = MarginInterface(self)
        self.margins.grid(row=5, column=0, pady=10, sticky="w")

        self.file_creation = ImagesFileCreation(self)
        self.file_creation.grid(row=6, column=0, columnspan=5, pady=10)

        # Enable row and column resizing
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def update_angles_needed(self, *args):
        # When you use trace_add method on a variable, it sends three additional arguments
        # to the callback function each time the variable’s value changes: variable_name, index_of_element, operation

        # When angles_needed is True, set best_orientation to False and disable its checkbox
        if self.angles_needed.get():
            self.best_orientation.set(False)
            self.page_orientation.disable_best_orientation()
        else:
            # Re-enable best_orientation checkbox when angles_needed is False
            self.page_orientation.enable_best_orientation()

    def update_best_orientation(self, *args):
        # When best_orientation is True, set angles_needed to False and disable its checkbox
        if self.best_orientation.get():
            self.angles_needed.set(False)
            self.input_interface.disable_angles_cb()
        else:
            # Re-enable angles_needed checkbox when best_orientation is False
            self.input_interface.enable_angles_cb()


class InputImagesInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # Configure grid layout to expand with resizing
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.create_frames()
        self.create_buttons()

    def disable_angles_cb(self):
        # Disable the Specify Angles checkbutton
        self.angles_cb.config(state="disabled")

    def enable_angles_cb(self):
        # Re-enable the Specify Angles checkbutton
        self.angles_cb.config(state="normal")

    def create_frames(self):
        # Frame to hold ScrollableCanvas

        # Scrollable canvas setup inside frame
        self.scrollable_canvas = ScrollableCanvas(self)
        self.scrollable_canvas.grid(row=0, column=1, sticky="nsew")

    def create_buttons(self):
        tk.Label(self, text="Images").grid(row=0, column=0, sticky="we")

        self.image_paths_button = tk.Button(self, text="Select Images", command=self.select_image_path)
        self.image_paths_button.grid(row=1, column=1, pady=10, sticky="w")

        self.angles_cb = tk.Checkbutton(self, text="Specify Angles", variable=self.parent.angles_needed,
                                        command=self.toggle_angles_option)

        self.angles_cb.grid(row=2, column=1, sticky="w")

    def select_image_path(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Image Files", "*.jpg *.jpeg *.png")])
        for file_path in file_paths:
            if file_path:
                self.parent.image_paths.append(file_path)
                self.display_image_list()

    def display_image_list(self):
        # Remove only widgets within `scrollable_content`, not the `scrollable_canvas` itself
        for widget in self.scrollable_canvas.scrollable_content.winfo_children():
            widget.destroy()

        self.parent.thumbnails.clear()

        for index, file_path in enumerate(self.parent.image_paths):
            img_title = os.path.basename(file_path)

            # Load image and create a thumbnail
            image = Image.open(file_path)
            image.thumbnail((50, 50))  # Resize image to 50x50 pixels
            thumbnail = ImageTk.PhotoImage(image)
            self.parent.thumbnails.append(thumbnail)  # Save reference to prevent garbage collection

            # Display thumbnail
            thumbnail_label = tk.Label(self.scrollable_canvas.scrollable_content, image=thumbnail)
            thumbnail_label.grid(row=index * 2, column=0, padx=5, pady=(5, 0))

            # Display image title below the thumbnail
            label = tk.Label(self.scrollable_canvas.scrollable_content, text=img_title, relief="ridge", padx=5, pady=5)
            label.grid(row=index * 2 + 1, column=0, padx=5, pady=5, sticky="nsew")

            delete_button = tk.Button(self.scrollable_canvas.scrollable_content, text="X",
                                      command=lambda idx=index: self.delete_image(idx))
            delete_button.grid(row=index * 2 + 1, column=1, padx=2)

            # Up and Down buttons to reorder images
            if index > 0:  # "Up" button is only enabled if it"s not the first item
                up_button = tk.Button(self.scrollable_canvas.scrollable_content, text="↑",
                                      command=lambda idx=index: self.move_image_up(idx))
                up_button.grid(row=index * 2 + 1, column=2, padx=2)

            if index < len(self.parent.image_paths) - 1:  # "Down" button is only enabled if it"s not the last item
                down_button = tk.Button(self.scrollable_canvas.scrollable_content, text="↓",
                                        command=lambda idx=index: self.move_image_down(idx))
                down_button.grid(row=index * 2 + 1, column=3, padx=2)

            if self.parent.angles_needed.get() and self.angles_cb.cget("state") != "disabled":
                angle_entry = tk.Entry(self.scrollable_canvas.scrollable_content, width=5)
                angle_entry.grid(row=index * 2 + 1, column=4, padx=2)
                self.parent.angle_entries[file_path] = angle_entry
                # Bind focus-out behavior to lose focus when clicked outside
                angle_entry.bind("<FocusOut>", lambda e: angle_entry.selection_clear())

        # Bind to release focus if clicked outside of an entry widget
        self.scrollable_canvas.bind("<Button-1>", lambda event: self.scrollable_canvas.focus_set())

    def move_image_up(self, index):
        if index > 0:
            self.parent.image_paths[index], self.parent.image_paths[index - 1] = self.parent.image_paths[index - 1], \
                                                                                 self.parent.image_paths[index]
            self.display_image_list()

    def move_image_down(self, index):
        if index < len(self.parent.image_paths) - 1:
            self.parent.image_paths[index], self.parent.image_paths[index + 1] = self.parent.image_paths[index + 1], \
                                                                                 self.parent.image_paths[index]
            self.display_image_list()

    def delete_image(self, index):
        del self.parent.image_paths[index]
        self.display_image_list()

    def toggle_angles_option(self):
        self.display_image_list()


class OutputOptionsInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        self.create_output_widgets()
        self.lay_out_output_widgets()

    def create_output_widgets(self):
        self.output_path_cb = tk.Checkbutton(self, text="Specify Output Path", variable=self.parent.output_path,
                                             command=self.toggle_output_path)
        self.output_path_label = tk.Label(self, text="No path selected", fg="grey")
        self.clear_output_button = tk.Button(self, text="Delete Path", command=self.clear_output_path)

    def lay_out_output_widgets(self):
        self.output_path_cb.grid(row=2, column=0, sticky="w")
        self.output_path_label.grid(row=2, column=1, columnspan=1, sticky="w")
        self.clear_output_button.grid(row=3, column=1, sticky="w")

    def toggle_output_path(self):
        if self.parent.output_path.get():
            self.select_output_path()
            self.output_path_cb.config(state="disabled")

    def select_output_path(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.parent.output_path_file = file_path
            self.output_path_label.config(text=self.parent.output_path_file, fg="black")

    def clear_output_path(self):
        self.parent.output_path.set(0)
        self.output_path_cb.config(state="active")
        self.output_path_label.config(text="No path selected", fg="grey")
        self.parent.output_path_file = None


class MultiplePagesInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        self.create_multiple_pages_widgets()
        self.lay_out_multiple_pages_widgets()

        self.toggle_grid_options()

        self.get_values()

    def create_multiple_pages_widgets(self):
        self.multiple_cb = tk.Checkbutton(self, text="Multiple Pages (Grid)", variable=self.parent.multiple_pages,
                                          command=self.toggle_grid_options)

        self.columns_label = tk.Label(self, text="Columns:")
        self.columns_entry = ttk.Spinbox(self, from_=1.0, to=16.0, textvariable=self.parent.columns_number, width=5,
                                         state="readonly")

        self.rows_label = tk.Label(self, text="Rows:")

        self.rows_entry = ttk.Spinbox(self, from_=1.0, to=16.0, textvariable=self.parent.rows_number, width=5,
                                      state="readonly")

    def lay_out_multiple_pages_widgets(self):
        self.multiple_cb.grid(row=3, column=0, sticky="w")

        self.columns_label.grid(row=3, column=1, sticky="e")
        self.columns_entry.grid(row=3, column=2)

        self.rows_label.grid(row=3, column=3, sticky="e")
        self.rows_entry.grid(row=3, column=4)

    def toggle_grid_options(self):
        state = "readonly" if self.parent.multiple_pages.get() else "disabled"
        self.columns_entry.config(state=state)
        self.rows_entry.config(state=state)

    def get_values(self):
        self.columns_number = self.columns_entry.get()
        self.rows_number = self.rows_entry.get()


class PagesOrientationInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        self.create_page_orientation_widgets()
        self.lay_out_page_orientation_widgets()
        self.lay_out_best_orientation()

    def create_page_orientation_widgets(self):
        tk.Label(self, text="Orientation:").grid(row=5, column=0, sticky="w")
        self.orientation_frame = tk.LabelFrame(self, text="Orientation")

        self.portrait_radio = tk.Radiobutton(self.orientation_frame, text="Portrait", variable=self.parent.orientation,
                                             value="portrait")
        self.landscape_radio = tk.Radiobutton(self.orientation_frame, text="Landscape",
                                              variable=self.parent.orientation, value="landscape")
        self.auto_radio = tk.Radiobutton(self.orientation_frame, text="Auto", variable=self.parent.orientation,
                                         value="auto", state="disabled")

    def lay_out_page_orientation_widgets(self):
        self.orientation_frame.grid(row=5, column=1, columnspan=4, sticky="w")

        self.portrait_radio.pack(anchor="w")
        self.landscape_radio.pack(anchor="w")
        self.auto_radio.pack(anchor="w")

    def lay_out_best_orientation(self):
        self.best_orientation_check = tk.Checkbutton(self, text="Best Orientation",
                                                     variable=self.parent.best_orientation,
                                                     command=self.toggle_best_orientation)
        self.best_orientation_check.grid(row=6, column=0, sticky="w")

    def toggle_best_orientation(self):

        if self.parent.best_orientation.get():
            # Disable the Specify Angles checkbox and enable "Auto" orientation
            self.auto_radio.config(state=tk.NORMAL)
            # Set orientation to "auto" by default when Best Orientation is enabled
            self.parent.orientation.set("auto")
        else:
            # Re-enable Specify Angles and disable "Auto" orientation
            self.auto_radio.config(state=tk.DISABLED)

            # Restore default orientation (keeping user’s previous selection)
            if self.parent.orientation.get() == "auto":
                self.parent.orientation.set("portrait")  # default to Portrait if "auto" was selected

    def disable_best_orientation(self):
        # Disable the Best Orientation checkbutton
        self.best_orientation_check.config(state="disabled")

    def enable_best_orientation(self):
        # Re-enable the Best Orientation checkbutton
        self.best_orientation_check.config(state="normal")


class MarginInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        tk.Label(self, text="Page Margin:").grid(row=7, column=0, sticky="w")

        self.page_margin_entry = tk.Entry(self, width=10, textvariable=self.parent.page_margin, validate="key",
                                          validatecommand=(self.register(MarginInterface.validate_int), "%P"))
        self.parent.page_margin.set(self.page_margin_entry.get())

        self.page_margin_entry.grid(row=7, column=1)

        tk.Label(self, text="Image Margin:").grid(row=7, column=2, sticky="w")

        self.image_margin_entry = tk.Entry(self, width=10, textvariable=self.parent.image_margin, validate="key",
                                           validatecommand=(self.register(MarginInterface.validate_int), "%P"))
        self.image_margin_entry.grid(row=7, column=3)
        self.parent.image_margin.set(self.image_margin_entry.get())

    @staticmethod
    def validate_int(value):
        if value.isdigit() or value == "":
            return True
        else:
            messagebox.showerror("Input Error", "Please enter a valid integer.")
            return False


class ImagesFileCreation(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        tk.Button(self, text="Generate PDF", command=self.generate_pdf).grid(row=0, column=0, columnspan=1, pady=10)

    def generate_pdf(self):
        # Retrieve values from GUI fields
        output_path = self.parent.output_path_file if self.parent.output_path.get() else None
        image_paths = self.parent.image_paths
        columns = int(self.parent.columns_number.get()) if self.parent.multiple_pages.get() else 1
        rows = int(self.parent.rows_number.get()) if self.parent.multiple_pages.get() else 1
        try:
            page_margin = int(self.parent.page_margin.get())
            image_margin = int(self.parent.image_margin.get())
        except tk.TclError:
            messagebox.showwarning("Warning!", "Margins are set to 0")
            page_margin, image_margin = 0, 0

        # If Best Orientation is selected, use that function
        if self.parent.best_orientation.get():
            self.generated_pdf = create_pdf_with_best_orientation_images(
                output_path=output_path,
                image_paths=image_paths,
                columns=columns,  # Columns and rows set to 1 as grid mode is not compatible
                rows=rows,
                orientation=self.parent.orientation.get(),
                page_margin=page_margin,
                image_margin=image_margin
            )
        else:
            # Otherwise, use the add_images_to_pdf_in_grid function
            angles = [int(entry.get()) if entry.get().isdigit() else 0 for entry in
                      self.parent.angle_entries.values()] if self.parent.angles_needed.get() else None
            self.generated_pdf = add_images_to_pdf_in_grid(
                output_path=output_path,
                image_paths=image_paths,
                columns=columns,
                rows=rows,
                angles=angles,
                orientation=self.parent.orientation.get(),
                page_margin=page_margin,
                image_margin=image_margin
            )

        # Optional feedback for success
        messagebox.showinfo("Success", "PDF generated successfully!")

        # Show Print button after generating PDF
        self.print_button = tk.Button(self, text="Print", command=self.open_print_window)
        self.print_button.grid(row=1, column=0, padx=10, pady=10)

    def open_print_window(self):
        PrintWindow(self, self.generated_pdf)


class ScrollableCanvas(tk.Frame):
    def __init__(self, parent, height=None):
        super().__init__(parent)
        # Canvas and scrollbars setup
        self.canvas = tk.Canvas(self, height=height, highlightthickness=0)
        self.vertical_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vertical_scrollbar.set)

        # Layout
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vertical_scrollbar.pack(side="right", fill="y")

        # Placeholder for the content frame (will be set by another class)
        self.scrollable_content = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_content, anchor="nw", tags="self.scrollable_content")

        # Bind resize events
        self.scrollable_content.bind("<Configure>", self.on_configure)

        # Bind mouse wheel scroll only when hovering over the scrollbars
        self.vertical_scrollbar.bind("<Enter>", self.enable_vertical_scroll)
        self.vertical_scrollbar.bind("<Leave>", self.disable_scroll)

        # Variables to track scroll direction and enablement
        self.scroll_direction = None
        self.vertical_scroll_enabled = False

    def on_configure(self, event=None):
        # Update the scroll region to match the size of the content frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        # Check if scrolling is needed based on content size vs. canvas size
        content_height = self.scrollable_content.winfo_reqheight()
        canvas_height = self.canvas.winfo_height()

        # Enable or disable vertical scrolling based on height
        self.vertical_scroll_enabled = content_height > canvas_height

    def enable_vertical_scroll(self, event):
        if self.vertical_scroll_enabled:
            self.scroll_direction = "vertical"
            self.bind_all("<MouseWheel>", self.on_mouse_wheel)

    def disable_scroll(self, event):
        self.scroll_direction = None
        self.unbind_all("<MouseWheel>")

    def on_mouse_wheel(self, event):
        if self.scroll_direction == "vertical":
            self.canvas.yview_scroll(-1 * (event.delta // 120), "units")


class PDFWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Files Configuration Window")
        self.iconbitmap(resource_path("printer.ico"))

        self.output_path = tk.BooleanVar()
        self.pdf_paths = []
        self.page_selections = None
        self.output_path_file = None
        self.pages_entries = {}

        self.input_interface = InputPDFInterface(self)
        self.input_interface.grid(row=0, column=0, columnspan=5, sticky="nsew")

        self.output_options = OutputOptionsInterface(self)
        self.output_options.grid(row=1, column=0, sticky="w")

        self.file_creation = PDFsFileCreation(self)
        self.file_creation.grid(row=2, column=0, columnspan=5, pady=10)

        # Enable row and column resizing
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)


class InputPDFInterface(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # Configure grid layout to expand with resizing
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.create_frames()
        self.create_buttons()

    def create_frames(self):
        self.scrollable_canvas = ScrollableCanvas(self)
        self.scrollable_canvas.grid(row=0, column=1, sticky="nsew")

    def create_buttons(self):
        tk.Label(self, text="PDF files").grid(row=0, column=0, sticky="w")

        self.image_paths_button = tk.Button(self, text="Select PDFs", command=self.select_pdf_path)
        self.image_paths_button.grid(row=1, column=1, pady=10, sticky="w")

    def select_pdf_path(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for file_path in file_paths:
            if file_path:
                self.parent.pdf_paths.append(file_path)
                self.display_pdf_list()

    def display_pdf_list(self):
        for widget in self.scrollable_canvas.scrollable_content.winfo_children():
            widget.destroy()

        for index, file_path in enumerate(self.parent.pdf_paths):
            pdf_title = os.path.basename(file_path)
            last_page = self.get_pdf_page_count(file_path)

            # Label for the PDF file name
            label = tk.Label(self.scrollable_canvas.scrollable_content, text=pdf_title, relief="ridge", padx=5, pady=5)
            label.grid(row=index * 2, column=1, padx=5, pady=5, sticky="w")

            delete_button = tk.Button(self.scrollable_canvas.scrollable_content, text="X",
                                      command=lambda idx=index: self.delete_pdf(idx))
            delete_button.grid(row=index * 2, column=2, padx=2)

            # Up and Down buttons to reorder images
            if index > 0:  # "Up" button is only enabled if it"s not the first item
                up_button = tk.Button(self.scrollable_canvas.scrollable_content, text="↑",
                                      command=lambda idx=index: self.move_pdf_up(idx))
                up_button.grid(row=index * 2, column=3, padx=2)

            if index < len(self.parent.pdf_paths) - 1:  # "Down" button is only enabled if it"s not the last item
                down_button = tk.Button(self.scrollable_canvas.scrollable_content, text="↓",
                                        command=lambda idx=index: self.move_pdf_down(idx))
                down_button.grid(row=index * 2, column=4, padx=2)

            pages_entry = tk.Entry(self.scrollable_canvas.scrollable_content, width=30)
            pages_entry.grid(row=index * 2, column=5, padx=2)

            pages_entry_example = tk.Label(self.scrollable_canvas.scrollable_content,
                                           text="Enter page ranges (e.g., 1-3,7,10-12):")
            pages_entry_example.grid(row=index * 2 + 1, column=5, padx=10)

            # Red warning label, initially hidden
            warning_label = tk.Label(self.scrollable_canvas.scrollable_content, text="", fg="red")
            warning_label.grid(row=index * 2, column=6, padx=5, pady=5, sticky="w")

            # Bind the Entry widget to trigger validation on key release
            pages_entry.bind("<KeyRelease>",
                             lambda event, entry=pages_entry, warning_label=warning_label, file_path=file_path,
                                    last_page=last_page:
                             self.validate_input(event, entry, warning_label, file_path, last_page))

            # Bind focus-out behavior to lose focus when clicked outside
            pages_entry.bind("<FocusOut>", lambda e: pages_entry.selection_clear())

        # Bind to release focus if clicked outside of an entry widget
        self.scrollable_canvas.bind("<Button-1>", lambda event: self.scrollable_canvas.focus_set())

    def get_pdf_page_count(self, pdf_path):
        try:
            pdf_document = pymupdf.open(pdf_path)
            page_count = pdf_document.page_count
            pdf_document.close()
            return page_count
        except Exception:
            return 0  # Return 0 if there was an error opening the file

    def validate_input(self, event, entry, warning_label, file_path, last_page):
        input_string = entry.get()
        result, error_message = InputPDFInterface.parse_page_ranges(input_string, last_page)
        if result is None:
            # Show warning label with the error message if validation fails
            warning_label.config(text=error_message)
        else:
            # Hide warning label if validation succeeds
            warning_label.config(text="")
        self.parent.pages_entries[file_path] = result

    @staticmethod
    def parse_page_ranges(input_str, last_page):
        # Initial validations
        if input_str.startswith("-") or "--" in input_str or "-," in input_str or ",-" in input_str:
            return None, "Invalid format: Cannot start with '-', and '-' or ',' cannot be adjacent."
        
        page_ranges = []
        if input_str == "":
            return page_ranges, ""
        
        parts = input_str.split(",")
        for part in parts:
            # Check for range (e.g., "1-3" or "10-")
            if "-" in part:
                start, end = part.split("-")
                # Validate the start and end page numbers
                if not start.isdigit():
                    return None, f"Invalid range start: '{start}' is not a number."
                start = int(start)
                # Handle open-ended range like "10-" as (10, last_page)
                if end == "":
                    end = last_page
                elif not end.isdigit():
                    return None, f"Invalid range end: '{end}' is not a number."
                else:
                    end = int(end)
                if start > end:
                    return None, f"Invalid range: Start page {start} cannot be greater than end page {end}."
                # Ensure pages don"t exceed the last page number
                if start > last_page or end > last_page:
                    return None, f"Invalid range: Pages cannot exceed the last page ({last_page})."
                page_ranges.append((start, end))
            else:
                # Validate single page entry
                if not part.isdigit():
                    return None, f"Invalid page: '{part}' is not a number."
                page_num = int(part)

                # Ensure single page doesn"t exceed the last page number
                if page_num > last_page:
                    return None, f"Invalid page: Page {page_num} exceeds the last page ({last_page})."
                page_ranges.append((page_num,))
        return page_ranges, ""  # Return the parsed ranges if valid

    def move_pdf_up(self, index):
        if index > 0:
            self.parent.pdf_paths[index], self.parent.pdf_paths[index - 1] = self.parent.pdf_paths[index - 1], \
                                                                             self.parent.pdf_paths[index]
            self.display_pdf_list()

    def move_pdf_down(self, index):
        if index < len(self.parent.pdf_paths) - 1:
            self.parent.pdf_paths[index], self.parent.pdf_paths[index + 1] = self.parent.pdf_paths[index + 1], \
                                                                             self.parent.pdf_paths[index]
            self.display_pdf_list()

    def delete_pdf(self, index):
        del self.parent.pdf_paths[index]
        self.display_pdf_list()


class PDFsFileCreation(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        tk.Button(self, text="Generate PDF", command=self.generate_pdf).grid(row=0, column=0, columnspan=5, pady=10)

    def generate_pdf(self):
        # Retrieve values from GUI fields
        output_path = self.parent.output_path_file if self.parent.output_path.get() else None
        pdf_paths = self.parent.pdf_paths
        page_selections = [entry for entry in
                           self.parent.pages_entries.values()] if not self.parent.page_selections else None

        self.generated_pdf = extract_and_merge_pdfs(pdf_paths, page_selections, output_path)

        # Optional feedback for success
        messagebox.showinfo("Success", "PDF generated successfully!")

        # Show Print button after generating PDF
        self.print_button = tk.Button(self, text="Print", command=self.open_print_window)
        self.print_button.grid(row=1, column=0, padx=10, pady=10)

    def open_print_window(self):
        PrintWindow(self, self.generated_pdf)


class PrintWindow(tk.Toplevel):
    def __init__(self, master, pdf_buffer_or_path):
        """
        Initialize the PrintWindow for configuring print options.

        :param master: Parent window
        :param pdf_buffer_or_path: PDF buffer or file path to print
        """
        super().__init__(master)
        self.title("Printer Settings")
        self.iconbitmap(resource_path("printer.ico"))

        self.pdf_buffer_or_path = pdf_buffer_or_path
        self.printer_name = None
        self.printer_manager = PrinterManager()
        self.temp_pdf_path = None

        # Create UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Printer selection
        tk.Label(self, text="List of available printers:").grid(row=0, column=0, padx=10, pady=5)
        printers = self.printer_manager.list_printers()
        self.printer_name = self.printer_manager.get_default_printer_name()
        self.printer_list = ttk.Combobox(self, values=printers, state="readonly")
        self.printer_list.current(printers.index(self.printer_name))
        self.printer_list.grid(row=0, column=1, padx=10, pady=5)
        self.printer_list.bind("<<ComboboxSelected>>", self.select_printer)

        # Copies
        tk.Label(self, text="Number of copies:").grid(row=1, column=0, padx=10, pady=5)
        self.copies_spinbox = tk.Spinbox(self, from_=1, to=100, width=5, command=self.check_sort_copies)
        self.copies_spinbox.grid(row=1, column=1, padx=10, pady=5)

        # Collate (sort Copies) (hidden initially)
        self.sort_copies = tk.BooleanVar(value=False)
        self.sort_copies_checkbox = tk.Checkbutton(self, text="Collate", variable=self.sort_copies)
        self.sort_copies_checkbox.grid(row=1, column=2, padx=10, pady=5)
        self.sort_copies_checkbox.grid_remove()  # Hide initially

        # Orientation
        tk.Label(self, text="Orientation:").grid(row=2, column=0, padx=10, pady=5)
        self.orientation = tk.IntVar(value=1)
        tk.Radiobutton(self, text="Portrait", variable=self.orientation, value=1).grid(row=2, column=1, sticky="w")
        tk.Radiobutton(self, text="Landscape", variable=self.orientation, value=2).grid(row=2, column=2, sticky="w")

        # Duplex printing
        tk.Label(self, text="Duplex:").grid(row=3, column=0, padx=10, pady=5)
        self.duplex = tk.BooleanVar(value=False)
        self.duplex_checkbox = tk.Checkbutton(self, text="Enable Duplex", variable=self.duplex,
                                               command=self.toggle_flip_side)
        self.duplex_checkbox.grid(row=3, column=1, sticky="w")

        # Flip side (hidden initially)
        self.flip_side_frame = tk.Frame(self)
        tk.Label(self.flip_side_frame, text="Flip Side:").grid(row=0, column=0, padx=10, pady=5)
        self.flip_side = tk.StringVar(value="long")
        tk.Radiobutton(self.flip_side_frame, text="Long Edge",
                       variable=self.flip_side, value="long").grid(row=0, column=1, sticky="w")
        tk.Radiobutton(self.flip_side_frame, text="Short Edge",
                       variable=self.flip_side, value="short").grid(row=0, column=2, sticky="w")

        # Only show flip side if duplex is enabled
        if self.duplex.get():
            self.flip_side_frame.grid(row=4, column=0, columnspan=3, pady=5)

        # Print Button
        tk.Button(self, text="Print", command=self.initiate_print).grid(row=5, column=0, columnspan=3, pady=10)

    def select_printer(self, event):
        self.printer_name = self.printer_list.get()

    def check_sort_copies(self):
        if int(self.copies_spinbox.get()) > 1:
            self.sort_copies_checkbox.grid()  # Show the checkbox
        else:
            self.sort_copies_checkbox.grid_remove()  # Hide the checkbox

    def create_temp_path(self):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(self.pdf_buffer_or_path.getvalue())
            self.temp_pdf_path = temp_pdf.name
            return self.temp_pdf_path

    def initiate_print(self):
        if not self.printer_name:
            print("Please select a printer.")
            messagebox.showwarning("Warning!", "Please select a printer.")
            return

        # Get user-selected settings
        try:
            copies = int(self.copies_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Please set the number of copies")
            return None
        orientation = self.orientation.get()
        duplex = self.duplex.get()
        flip_side = self.flip_side.get()
        sort_copies = self.sort_copies.get()

        self.printer_manager = PrinterManager(self.printer_name)
        # Call the print function with parameters
        if isinstance(self.pdf_buffer_or_path, io.BytesIO):
            self.create_temp_path()
            messagebox.showinfo("Success", "Document is sent to printer spooler successfully! \
            Please, don't close the windows until the next messagebox will appear")
            try:
                self.printer_manager.print_pdf(self.temp_pdf_path, copies, orientation, duplex, flip_side, sort_copies)
                messagebox.showinfo("Success", "The file is on printer. Wait the printer will finish its job. Good luck!")
            except Exception as exception:
                messagebox.showerror("Error", f"An error occurs during printing: \"{exception.strerror}\"! Please, check printer setting or select another printer")
            finally:
                os.remove(self.temp_pdf_path)
        else:
            self.printer_manager.print_pdf(self.pdf_buffer_or_path, copies, orientation, duplex, flip_side, sort_copies)
            messagebox.showinfo("Success", "The file is on printer. Wait the printer will finish its job. Good luck!")

    def toggle_flip_side(self):
        if self.duplex.get():
            self.flip_side_frame.grid(row=4, column=0, columnspan=3, pady=5)
        else:
            self.flip_side_frame.grid_remove()


if __name__ == "__main__":
    app = App()
