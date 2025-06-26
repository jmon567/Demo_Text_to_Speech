'''
Run this in the command line first

pip install PyMuPDF mutagen docx2txt
'''

import tkinter as tk
from tkinter import filedialog
from gtts import gTTS
import fitz  # PyMuPDF
from mutagen.mp3 import MP3
import threading
import docx2txt  # For Word .doc and .docx support

# Define the main app class
class TalkyFiles(tk.Tk):
    def __init__(self):
        super().__init__()

        # Improved UX: Set window title and size
        self.title('TalkyFiles - Convert to Audio')
        self.geometry('400x250')  # Enlarged slightly to accommodate more info

        # Create and place widgets
        self.create_widgets()

    def create_widgets(self):
        # Label for instructions
        self.label = tk.Label(self, text="Select a PDF, Text, or Word file to convert to audio")
        self.label.pack(pady=10)

        # Button to select file
        self.select_button = tk.Button(self, text="Select File", command=self.select_file)
        self.select_button.pack(pady=5)

        # Label to display conversion progress 
        self.progress_label = tk.Label(self, text="")
        self.progress_label.pack(pady=5)

    def select_file(self):
        # File dialog to select PDF, TXT, or DOC/DOCX
        filetypes = [
            ("PDF files", "*.pdf"), 
            ("Text files", "*.txt"),
            ("Word Documents", "*.doc;*.docx")
        ]
        self.file_path = filedialog.askopenfilename(defaultextension=".pdf", filetypes=filetypes)

        if self.file_path:
            self.select_button['state'] = 'disabled'  # Disable button during conversion
            self.thread = threading.Thread(target=self.convert_to_audio)
            self.thread.start()

    def convert_to_audio(self):
        try:
            self.progress_label['text'] = "Converting file to audio..."

            filename, extension = self.file_path.rsplit('.', 1)  # Get filename and extension

            if extension.lower() == 'pdf':
                # PDF conversion process (same as before)
                doc = fitz.open(self.file_path)
                text = ""
                for page_number in range(len(doc)):
                    page = doc.load_page(page_number)
                    text += page.get_text()
                doc.close()

            elif extension.lower() == 'txt':
                # Text file conversion
                with open(self.file_path, 'r') as file:
                    text = file.read()

            elif extension.lower() in ['doc', 'docx']:
                # Word document conversion
                text = docx2txt.process(self.file_path)

            else:
                raise ValueError("Unsupported file type. Please select a PDF, TXT, or Word file.")

            tts = gTTS(text, lang='en')  
            audio_file = filename + '.mp3' 
            tts.save(audio_file)

            audio = MP3(audio_file)
            audio_length = audio.info.length

            # Success message now includes file type
            self.progress_label['text'] = f"Conversion Complete!\nFile Type: {extension.upper()}\nAudio File: {audio_file}\nRuntime: {audio_length:.2f} seconds"

        except Exception as e:
            self.progress_label['text'] = f"Conversion Failed: {e}"
        finally:
            self.select_button['state'] = 'normal'  # Re-enable button 

# Run the app
if __name__ == "__main__":
    app = TalkyFiles()
    app.mainloop()
