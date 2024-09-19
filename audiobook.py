'''
The code is somewhat buggy and often may not work as intended. I only made this so I can dictate a nice narration of the story 
I am currently working on. I usually do not exceed narrating any more than 50 pages at a time and even then it tends to
overshoot or undershoot. Depending on the PC setup one has it can also take a considerable amount of time for conversion
to complete. Use it at your own risk if you must.

'''

import os
import re
import time
from pdfreader import SimplePDFViewer, PageDoesNotExist
from gtts import gTTS
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def clean_text(text):
    
    # Expressions to remove unwanted non-text content as in images.
    text = re.sub(r'(?i)page\s*\d+|[-—–]\s*\d+\s*[-—–]|\(\d+\)|^\d+\s*$', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def pdf_to_audiobook(pdf_path, output_path, start_page, end_page, lang='en', speed=1.0, progress_bar=None):
    try:
        text = ""
        start_time = time.time()

        # Open the PDF and set up the viewer
        with open(pdf_path, 'rb') as pdf_file:
            viewer = SimplePDFViewer(pdf_file)
            total_pages = end_page - start_page + 1

            # Ensure the pages are within bounds
            viewer.navigate(1)
            viewer.render()
            max_page_count = viewer.canvas.num_pages
            if start_page < 1 or end_page > max_page_count:
                messagebox.showerror("Error", f"Invalid page range. PDF has {max_page_count} pages.")
                return

            # Extract text from the specified range
            for page_num in range(start_page, end_page + 1):
                try:
                    viewer.navigate(page_num)
                    viewer.render()

                    # Extract text elements from the page
                    page_text_elements = viewer.canvas.strings
                    if not page_text_elements:
                        continue

                    # Combine the extracted text
                    page_text = ''.join(page_text_elements)
                    cleaned_text = clean_text(page_text)
                    
                    if cleaned_text.strip():
                        text += cleaned_text + " "

                    # Update progress bar if extract is available for conversion
                    if progress_bar:
                        progress_bar['value'] = ((page_num - start_page + 1) / total_pages) * 100
                        progress_bar.update()

                except PageDoesNotExist:
                    messagebox.showwarning("Warning", f"Page {page_num} does not exist.")
                    continue

        # Calculate and show time taken for extraction
        extraction_time = time.time() - start_time
        estimated_time = extraction_time / total_pages * max_page_count
        messagebox.showinfo("Info", f"Text extraction completed in {extraction_time:.2f} seconds. Estimated total time for entire document: {estimated_time:.2f} seconds.")

        # Adjust speech speed: gTTS only supports slow or normal speeds. I set the default at 1.0 : My experience says 3 - 4 times that is preferable
        slow = speed < 1

        if text.strip():  # To ensure that there's text to convert
            tts = gTTS(text=text, lang=lang, slow=slow)
            tts.save(output_path)
            messagebox.showinfo("Success", f"Audiobook saved to {output_path}")
        else:
            messagebox.showinfo("Info", "No text found in the selected page range.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")], title="Select a PDF File"
    )
    pdf_path_var.set(file_path)

def browse_output():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".mp3", filetypes=[("MP3 files", "*.mp3")], title="Save As"
    )
    output_path_var.set(file_path)

def start_conversion():
    pdf_path = pdf_path_var.get()
    output_path = output_path_var.get()
    try:
        start_page = int(start_page_var.get())
        end_page = int(end_page_var.get())
        speed = float(speed_var.get())

        if not os.path.isfile(pdf_path):
            raise ValueError("Please select a valid PDF file.")
        if not output_path:
            raise ValueError("Please select an output file path.")
        if start_page < 1 or end_page < start_page:
            raise ValueError("Please enter a valid page range.")

        # Start the conversion with progress bar
        pdf_to_audiobook(pdf_path, output_path, start_page, end_page, speed=speed, progress_bar=progress)

    except ValueError as ve:
        messagebox.showerror("Input Error", str(ve))
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Set up the GUI
root = tk.Tk()
root.title("PDF to Audiobook Converter")

# Variables for storing file paths and settings. Doing it all in the same directory would be preferable
pdf_path_var = tk.StringVar()
output_path_var = tk.StringVar()
start_page_var = tk.StringVar(value='1')
end_page_var = tk.StringVar(value='1')
speed_var = tk.StringVar(value='1.0')

# GUI Layout
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

ttk.Label(frame, text="PDF File:").grid(row=0, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=pdf_path_var, width=50).grid(row=0, column=1)
ttk.Button(frame, text="Browse", command=browse_pdf).grid(row=0, column=2)

ttk.Label(frame, text="Output MP3:").grid(row=1, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=output_path_var, width=50).grid(row=1, column=1)
ttk.Button(frame, text="Save As", command=browse_output).grid(row=1, column=2)

ttk.Label(frame, text="Start Page:").grid(row=2, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=start_page_var, width=10).grid(row=2, column=1, sticky=tk.W)

ttk.Label(frame, text="End Page:").grid(row=3, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=end_page_var, width=10).grid(row=3, column=1, sticky=tk.W)

ttk.Label(frame, text="Speed (0.5 - 2.0):").grid(row=4, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=speed_var, width=10).grid(row=4, column=1, sticky=tk.W)

ttk.Button(frame, text="Convert to Audiobook", command=start_conversion).grid(row=5, column=0, columnspan=3)

# Add progress bar for conversion, sit back and relax. It usually takes about 10-20 mins for 50 pages, then again my PC is low end.
progress = ttk.Progressbar(frame, orient='horizontal', mode='determinate', length=400)
progress.grid(row=6, column=0, columnspan=3, pady=10)

root.mainloop()
