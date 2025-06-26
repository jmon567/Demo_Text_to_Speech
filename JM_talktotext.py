

import tkinter as tk
from tkinter import filedialog, ttk  # Import ttk for themed widgets
from gtts import gTTS
import fitz  # PyMuPDF for PDF processing
from mutagen.mp3 import MP3  # To get audio file length
import threading  # For running conversion in a separate thread
import docx2txt  # For Word .doc and .docx support
import os # For path manipulation

# Define the main application class, inheriting from ttk.Frame for a themed container
class TalkyFiles(tk.Tk):
    """
    TalkyFiles Application: Converts PDF, TXT, and Word files to audio (MP3) using gTTS.
    Features an enhanced user interface with ttk widgets, dynamic feedback, and threading
    to keep the UI responsive during file conversion.
    """
    def __init__(self):
        """
        Initializes the TalkyFiles application window.
        Sets up the main window properties and calls the widget creation method.
        """
        super().__init__()

        # --- Window Configuration ---
        self.title('TalkyFiles - Convert Document to Audio') # Set the window title
        self.geometry('550x300') # Set a slightly larger initial window size
        self.resizable(False, False) # Make the window non-resizable for a consistent layout

        # Apply a modern Tkinter theme. 'clam' is often a good choice for a clean look.
        # Other themes include 'alt', 'default', 'vista', 'xpnative', 'classic'.
        style = ttk.Style(self)
        style.theme_use('clam') # Use 'clam' theme for a modern appearance

        # Configure styles for specific widgets for a more polished look
        style.configure('TFrame', background='#f0f0f0') # Light gray background for frames
        style.configure('TLabel', background='#f0f0f0', font=('Inter', 11)) # Label font and background
        style.configure('TButton', font=('Inter', 10, 'bold'), padding=8) # Button font and padding
        style.map('TButton',
                   foreground=[('pressed', 'red'), ('active', 'blue')], # Change text color on press/hover
                   background=[('pressed', '!focus', 'SystemButtonText'), ('active', '#e0e0e0')]) # Change background on press/hover

        # --- State Variables ---
        # StringVar to hold the selected file path for display
        self.file_path_var = tk.StringVar(self, value="No file selected")
        # StringVar to display progress and status messages
        self.progress_var = tk.StringVar(self, value="Select a PDF, Text, or Word file to convert.")
        # Variable to hold the path of the converted audio file for display
        self.audio_output_var = tk.StringVar(self, value="")
        # Variable to hold the runtime of the audio file for display
        self.audio_runtime_var = tk.StringVar(self, value="")

        # --- Create and Place Widgets ---
        self.create_widgets()

    def create_widgets(self):
        """
        Creates and arranges all GUI widgets within the application window.
        Uses a grid layout for flexible and organized placement.
        """
        # Create a main frame to contain all widgets, providing consistent padding
        main_frame = ttk.Frame(self, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True) # Make the frame fill and expand with the window

        # --- Instructions Label ---
        # A main instruction label guiding the user
        ttk.Label(main_frame, textvariable=self.progress_var, wraplength=450, justify=tk.CENTER).grid(
            row=0, column=0, columnspan=3, pady=(0, 15), sticky='ew'
        )

        # --- File Selection Section ---
        # Label to indicate where the selected file path will be displayed
        ttk.Label(main_frame, text="Selected File:").grid(row=1, column=0, sticky='w', pady=(5, 5))
        # Label to display the chosen file's path (updated via file_path_var)
        ttk.Label(main_frame, textvariable=self.file_path_var, font=('Inter', 9, 'italic'), foreground='gray').grid(
            row=1, column=1, columnspan=2, sticky='ew', pady=(5, 5)
        )

        # Button to trigger the file selection dialog
        self.select_button = ttk.Button(main_frame, text="Select File", command=self.select_file)
        self.select_button.grid(row=2, column=0, columnspan=3, pady=(10, 20)) # Place in the center

        # --- Output Information Section ---
        # Label to show the output audio file name
        ttk.Label(main_frame, text="Audio Output:").grid(row=3, column=0, sticky='w', pady=(5, 0))
        ttk.Label(main_frame, textvariable=self.audio_output_var, font=('Inter', 10, 'bold')).grid(
            row=3, column=1, columnspan=2, sticky='ew', pady=(5, 0)
        )

        # Label to show the estimated audio runtime
        ttk.Label(main_frame, text="Estimated Runtime:").grid(row=4, column=0, sticky='w', pady=(0, 5))
        ttk.Label(main_frame, textvariable=self.audio_runtime_var, font=('Inter', 10)).grid(
            row=4, column=1, columnspan=2, sticky='ew', pady=(0, 5)
        )

        # --- Grid Column Configuration ---
        # Configure columns to expand proportionally, making the layout responsive
        main_frame.grid_columnconfigure(0, weight=0) # Column 0 (labels) won't expand
        main_frame.grid_columnconfigure(1, weight=1) # Column 1 (content) will expand
        main_frame.grid_columnconfigure(2, weight=0) # Column 2 (empty/padding) won't expand


    def select_file(self):
        """
        Opens a file dialog for the user to select a PDF, TXT, or Word document.
        If a file is selected, it updates the UI and starts the conversion process
        in a separate thread to prevent the UI from freezing.
        """
        filetypes = [
            ("PDF files", "*.pdf"),
            ("Text files", "*.txt"),
            ("Word Documents", "*.doc;*.docx"),
            ("All supported files", "*.pdf;*.txt;*.doc;*.docx") # Combined option
        ]
        # Open the file dialog and get the selected file path
        self.file_path = filedialog.askopenfilename(
            title="Select Document File",
            defaultextension=".pdf",
            filetypes=filetypes
        )

        if self.file_path:
            # Update the selected file display
            self.file_path_var.set(os.path.basename(self.file_path))
            # Disable the select button to prevent multiple conversions simultaneously
            self.select_button['state'] = 'disabled'
            # Update the progress message
            self.progress_var.set("File selected. Starting conversion...")
            # Clear previous output details
            self.audio_output_var.set("")
            self.audio_runtime_var.set("")

            # Start the conversion in a new thread
            # This is crucial for keeping the GUI responsive during long operations
            self.conversion_thread = threading.Thread(target=self.convert_to_audio)
            self.conversion_thread.start()
        else:
            # If no file was selected, reset the status message
            self.progress_var.set("File selection cancelled. Select a file to convert.")
            self.file_path_var.set("No file selected")


    def convert_to_audio(self):
        """
        Handles the core logic of converting the selected document to audio.
        This method runs in a separate thread. It extracts text based on file type,
        uses gTTS to generate speech, saves it as an MP3, and then updates the UI
        upon completion or error.
        """
        try:
            # Indicate that conversion is in progress
            self.progress_var.set("Converting file to audio... Please wait.")

            # Get filename and extension for appropriate processing
            filename_with_ext = os.path.basename(self.file_path)
            base_filename, extension = os.path.splitext(filename_with_ext)
            extension = extension.lower().lstrip('.') # Normalize extension (e.g., '.pdf' -> 'pdf')

            extracted_text = ""
            if extension == 'pdf':
                # --- PDF Text Extraction (using PyMuPDF/fitz) ---
                doc = fitz.open(self.file_path)
                for page_number in range(len(doc)):
                    page = doc.load_page(page_number)
                    extracted_text += page.get_text()
                doc.close()
                file_type_display = "PDF"

            elif extension == 'txt':
                # --- Text File Extraction ---
                with open(self.file_path, 'r', encoding='utf-8') as file: # Specify encoding for broader compatibility
                    extracted_text = file.read()
                file_type_display = "Text File"

            elif extension in ['doc', 'docx']:
                # --- Word Document Extraction (using docx2txt) ---
                extracted_text = docx2txt.process(self.file_path)
                file_type_display = "Word Document"

            else:
                # Handle unsupported file types
                raise ValueError("Unsupported file type. Please select a PDF, TXT, or Word file.")

            # Ensure some text was extracted before proceeding
            if not extracted_text.strip():
                raise ValueError("No readable text found in the selected file.")

            # --- Text-to-Speech Conversion (using gTTS) ---
            # Create a gTTS object. 'lang' specifies the language (e.g., 'en' for English).
            tts = gTTS(text=extracted_text, lang='en', slow=False) # slow=False for normal speed
            # Define the output audio file name
            audio_file_name = f"{base_filename}_audio.mp3" # Appending _audio to avoid overwriting original
            # Construct the full path for the audio file (saves in the same directory as original file)
            audio_file_path = os.path.join(os.path.dirname(self.file_path), audio_file_name)
            tts.save(audio_file_path)

            # --- Get Audio File Length (using mutagen.mp3) ---
            # Open the saved MP3 file to get its duration
            audio = MP3(audio_file_path)
            audio_length = audio.info.length # Length in seconds

            # --- Update UI on Success (using self.after to run on main thread) ---
            # Schedule a function call on the main Tkinter thread to update the UI
            self.after(0, lambda: self.update_ui_on_success(
                file_type_display, audio_file_path, audio_length
            ))

        except Exception as e:
            # --- Update UI on Failure (using self.after to run on main thread) ---
            self.after(0, lambda: self.update_ui_on_failure(e))
        finally:
            # --- Re-enable Button (using self.after to run on main thread) ---
            # Ensure the select button is re-enabled whether conversion succeeds or fails
            self.after(0, lambda: self.select_button.config(state='normal'))

    def update_ui_on_success(self, file_type, audio_file_path, audio_length):
        """
        Updates the UI with success messages and audio file details after conversion.
        This method is called from the main thread via self.after().
        """
        # Set success message
        self.progress_var.set(f"Conversion Complete! Your audio file is ready.")
        # Display the converted audio file name
        self.audio_output_var.set(f"{os.path.basename(audio_file_path)}")
        # Display the audio runtime in a user-friendly format
        self.audio_runtime_var.set(f"{audio_length:.2f} seconds ({audio_length / 60:.1f} minutes)")
        # Optionally, provide a visual cue like changing label color
        self.progress_var.trace_add('write', lambda *args: self.set_label_color(self.progress_var, 'green'))
        self.after(3000, lambda: self.set_label_color(self.progress_var, None)) # Reset color after 3 seconds

    def update_ui_on_failure(self, error):
        """
        Updates the UI with an error message if conversion fails.
        This method is called from the main thread via self.after().
        """
        self.progress_var.set(f"Conversion Failed: {error}")
        self.audio_output_var.set("N/A")
        self.audio_runtime_var.set("N/A")
        # Optionally, provide a visual cue like changing label color to red for errors
        self.progress_var.trace_add('write', lambda *args: self.set_label_color(self.progress_var, 'red'))
        self.after(3000, lambda: self.set_label_color(self.progress_var, None)) # Reset color after 3 seconds

    def set_label_color(self, string_var, color):
        """Helper to change a label's foreground color dynamically."""
        # Find the label widget associated with the given StringVar
        # This is a bit of a hacky way to get the widget, but works for simple cases.
        # A more robust solution would be to pass the widget itself or store references.
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Label) and child.cget('textvariable') == string_var._name:
                        child.config(foreground=color)
                        break

# --- Main Application Entry Point ---
if __name__ == "__main__":
    app = TalkyFiles() # Create an instance of the application
    app.mainloop() # Start the Tkinter event loop, which runs the application
