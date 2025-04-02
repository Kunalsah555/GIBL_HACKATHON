import cv2
import numpy as np
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import os
import logging
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from ttkthemes import ThemedTk
import threading
from tkinter import ttk
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Ensure that Tesseract path is set up correctly
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
poppler_path = r"C:\Users\HP\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# Preprocessing function with adaptive thresholding
def preprocess_image(img):
    try:
        open_cv_image = np.array(img)
        gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        _, thresholded = cv2.threshold(blurred, 150, 255, cv2.THRESH_BINARY)
        return thresholded
    except Exception as e:
        logging.error(f"Error in preprocessing: {e}")
        return img

# Function to save text to a file
def save_text_to_file(text, file_name):
    try:
        with open(file_name + ".txt", "w", encoding="utf-8") as txt_file:
            txt_file.write(text)
        with open(file_name + ".json", "w", encoding="utf-8") as json_file:
            json.dump({"extracted_text": text}, json_file, indent=4)
        logging.info(f"Text saved successfully to {file_name}.txt and {file_name}.json")
    except Exception as e:
        logging.error(f"Error saving text to file: {e}")

# Function to extract text from a single image
def extract_text_from_image(image_path):
    try:
        if not os.path.exists(image_path):
            logging.error(f"File not found: {image_path}")
            return "Error: Image file not found."
        
        img = Image.open(image_path)
        processed_image = preprocess_image(img)
        pil_image = Image.fromarray(processed_image)
        text = pytesseract.image_to_string(pil_image).strip()
        
        save_text_to_file(text, "image_text")
        return text
    except Exception as e:
        logging.error(f"Error extracting text from image: {e}")
        return "Error: Unable to extract text."

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    try:
        if not os.path.exists(pdf_path):
            logging.error(f"File not found: {pdf_path}")
            return "Error: PDF file not found."
        
        images = convert_from_path(pdf_path, poppler_path=poppler_path)
        extracted_text = ""
        
        # Set progress bar to determinate mode
        progress_bar.config(mode='determinate', maximum=len(images))
        progress_bar.start()

        for i, img in enumerate(images):
            processed_image = preprocess_image(img)
            pil_image = Image.fromarray(processed_image)
            text = pytesseract.image_to_string(pil_image)
            extracted_text += f"Page {i + 1}:\n{text}\n\n"

            # Update progress bar
            progress_bar['value'] = i + 1
            root.update_idletasks()  # Update UI

        extracted_text = extracted_text.strip()
        save_text_to_file(extracted_text, "pdf_text")
        progress_bar.stop()
        return extracted_text
    except Exception as e:
        logging.error(f"Error extracting text from PDF: {e}")
        return "Error: Unable to extract text from PDF."

# Function to handle file type and call the correct extractor
def extract_text(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == ".pdf":
        return extract_text_from_pdf(file_path)
    elif file_extension in [".jpg", ".jpeg", ".png"]:
        return extract_text_from_image(file_path)
    else:
        return "Error: Unsupported file format."

# Cloud integration (Google Drive example)
def upload_to_google_drive(file_path):
    def upload():
        try:
            gauth = GoogleAuth()
            gauth.LocalWebserverAuth()  # This creates a local webserver and auto handles authentication.
            drive = GoogleDrive(gauth)

            file_drive = drive.CreateFile({'title': os.path.basename(file_path)})
            file_drive.Upload()
            logging.info(f"File uploaded to Google Drive: {file_path}")
            messagebox.showinfo("Success", f"File uploaded to Google Drive: {file_path}")
        except Exception as e:
            logging.error(f"Error uploading file to Google Drive: {e}")
            messagebox.showerror("Error", f"Error uploading file: {e}")

    threading.Thread(target=upload).start()  # Run upload in background thread

# Function to select a file using a file dialog
def browse_file():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf"), ("Image files", "*.jpg;*.jpeg;*.png")])
    if file_paths:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, ', '.join(file_paths))
        label_file_format.config(text=f"Detected File Format: {', '.join([os.path.splitext(f)[1].upper() for f in file_paths])}")

# Function to extract and display text
def extract_and_display_text():
    file_paths = entry_file_path.get().split(', ')  # Split comma-separated file paths
    if not file_paths:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        # Start the progress bar animation
        progress_bar.start()

        # Update the status message to indicate extraction in progress
        status_label.config(text="Extracting text... Please wait...")

        all_text = ""
        for file_path in file_paths:
            extracted_text = extract_text(file_path)
            all_text += extracted_text + "\n\n"  # Combine all extracted texts

        # Stop the progress bar animation when OCR is complete
        progress_bar.stop()

        # Update the status message to indicate completion
        status_label.config(text="Extraction complete.")

        # Display the combined extracted text
        text_output.delete(1.0, tk.END)  # Clear previous output
        text_output.insert(tk.END, all_text)  # Insert combined text

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        progress_bar.stop()
        status_label.config(text="Error during extraction.")

# Function to save the extracted text to a file
def save_extracted_text():
    extracted_text = text_output.get(1.0, tk.END).strip()
    if not extracted_text:
        messagebox.showwarning("Warning", "No text to save.")
        return

    # Ask the user for the file name and location to save
    save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt"), ("JSON Files", "*.json")])
    if save_path:
        try:
            # Save as text and JSON files
            save_text_to_file(extracted_text, os.path.splitext(save_path)[0])
            messagebox.showinfo("Success", f"Text saved successfully to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving text: {e}")

# Function for handling background threading for OCR extraction
def run_ocr_in_background():
    file_paths = entry_file_path.get().split(', ')  # Split comma-separated file paths
    if not file_paths:
        messagebox.showerror("Error", "Please select a file first.")
        return

    # Start the extraction in a background thread
    threading.Thread(target=extract_and_display_text).start()

# Set up the main Tkinter window (Themed for better look)
root = ThemedTk()
root.title("Document OCR and Text Extraction")
root.set_theme("arc")  # Use a pre-defined theme

# Frame for file selection
frame_file = tk.Frame(root)
frame_file.pack(padx=20, pady=10)

label_file = tk.Label(frame_file, text="Select PDF/Image file:", font=("Arial", 12), fg="black")
label_file.pack(side=tk.LEFT)

entry_file_path = tk.Entry(frame_file, width=50, font=("Arial", 12), relief="solid")
entry_file_path.pack(side=tk.LEFT)

button_browse = tk.Button(frame_file, text="Browse", font=("Arial", 12), command=browse_file, bg="#4CAF50", fg="white")
button_browse.pack(side=tk.LEFT, padx=5)

# Label to display detected file format
label_file_format = tk.Label(root, text="Detected File Format: None", font=("Arial", 12), fg="gray")
label_file_format.pack(padx=10, pady=5)

# Frame for extracting and displaying text
frame_output = tk.Frame(root)
frame_output.pack(padx=20, pady=10)

button_extract = tk.Button(frame_output, text="Extract Text", font=("Arial", 12), command=run_ocr_in_background, bg="#4CAF50", fg="white")
button_extract.pack(pady=10)

# Scrolled text widget for displaying the extracted text
text_output = ScrolledText(root, width=70, height=15, font=("Arial", 12))
text_output.pack(padx=20, pady=10)

# Frame for saving extracted text
frame_save = tk.Frame(root)
frame_save.pack(padx=20, pady=10)

button_save = tk.Button(frame_save, text="Save Text", font=("Arial", 12), command=save_extracted_text, bg="#2196F3", fg="white")
button_save.pack()

# Frame for cloud upload
frame_upload = tk.Frame(root)
frame_upload.pack(padx=20, pady=10)

button_upload = tk.Button(frame_upload, text="Upload to Google Drive", font=("Arial", 12), command=lambda: upload_to_google_drive(entry_file_path.get()), bg="#FF5722", fg="white")
button_upload.pack()

# Add a progress bar for OCR extraction
progress_frame = tk.Frame(root)
progress_frame.pack(padx=20, pady=10)

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=500, mode="indeterminate")
progress_bar.pack(pady=5)

# Status label to show the current status of extraction/upload
status_label = tk.Label(root, text="Ready", font=("Arial", 12), fg="green")
status_label.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
