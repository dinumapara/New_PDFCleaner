import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

# Import OCR and image conversion libraries
from pdf2image import convert_from_path
import pytesseract

import logging
import tempfile
from typing import Optional

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')


def is_blank_page(page, file_path: str, page_num: int, poppler_path: Optional[str] = None) -> bool:
    """
    Determine if a page is blank.

    1. Attempt to extract text with PyPDF2 and remove whitespace.
       If any meaningful text remains, the page is non-blank.
    2. If the extracted text is empty, convert the page to an image and run OCR.
    3. If the OCR output, after stripping, is shorter than OCR_TEXT_LENGTH_THRESHOLD,
       treat the page as blank.

    Parameters:
      - page: A PyPDF2 page object.
      - file_path: The path to the original PDF.
      - page_num: The current page number (1-indexed).
      - poppler_path: Optional path to the Poppler binaries (if required on your system).

    Returns:
      True if the page is considered blank; otherwise False.
    """
    text = page.extract_text()
    if text is not None and text.strip() != "":
        return False  # Page has meaningful text

    # Use OCR if text extraction yields nothing (or only whitespace)
    try:
        if poppler_path:
            images = convert_from_path(file_path, first_page=page_num, last_page=page_num, poppler_path=poppler_path)
        else:
            images = convert_from_path(file_path, first_page=page_num, last_page=page_num)
        if images:
            image = images[0]
            ocr_text = pytesseract.image_to_string(image)
            if ocr_text and len(ocr_text.strip()) >= OCR_TEXT_LENGTH_THRESHOLD:
                return False
            else:
                return True
        else:
            return True
    except Exception as e:
        logging.error(f"OCR failed for page {page_num}: {e}")
        # If OCR fails, to be safe, consider the page blank.
        return True


def process_pdf(file_path: str, poppler_path: Optional[str] = None) -> None:
    """
    Process the given PDF:
      - Make a backup copy.
      - Remove blank pages based on text extraction and OCR.
      - Save the modified PDF (overwriting the original file).
      - Delete the backup copy if successful.
    """
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    backup_path = os.path.join(dir_name, base_name + ".backup")
    try:
        shutil.copy(file_path, backup_path)
        logging.info(f"Backup created: {backup_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create backup copy:\n{e}")
        return
    try:
        reader = PdfReader(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read PDF:\n{e}")
        return
    writer = PdfWriter()
    removed_pages = 0
    total_pages = len(reader.pages)
    # Process each page: add non-blank pages to the writer.
    for i, page in enumerate(reader.pages, start=1):
        if is_blank_page(page, file_path, i, poppler_path):
            removed_pages += 1
            logging.info(f"Page {i} is blank and will be removed.")
        else:
            writer.add_page(page)
    if len(writer.pages) == 0:
        messagebox.showwarning("Warning", "All pages are blank. The PDF will not be modified.")
        return
    # Write the modified PDF to a secure temporary file.
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            temp_file = tmp.name
        with open(temp_file, "wb") as f_out:
            writer.write(f_out)
        logging.info("Modified PDF written to temporary file.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write modified PDF:\n{e}")
        return
    # Replace the original file with the modified file.
    try:
        shutil.move(temp_file, file_path)
        logging.info("Original PDF replaced with modified PDF.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to replace the original PDF:\n{e}")
        return
    # Delete the backup copy since the process finished successfully.
    try:
        os.remove(backup_path)
        logging.info("Backup copy deleted.")
    except Exception as e:
        logging.error(f"Could not remove backup copy: {e}")
    messagebox.showinfo("Success",
                        f"Modified PDF saved.\nTotal pages: {total_pages}\nRemoved blank pages: {removed_pages}")


def process_folder(poppler_path: Optional[str] = None) -> None:
    """
    Process all PDF files in a selected folder.
    """
    folder_path = filedialog.askdirectory(title="Select a Folder containing PDF files")
    if not folder_path:
        messagebox.showinfo("Info", "No folder selected.")
        return
    pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    if not pdf_files:
        messagebox.showinfo("Info", "No PDF files found in the selected folder.")
        return
    processed_count = 0
    for file_path in pdf_files:
        process_pdf(file_path, poppler_path)
        processed_count += 1
    messagebox.showinfo("Success", f"Processed {processed_count} PDF files.")


def main():
    # Create a simple Tkinter file dialog to choose the PDF file or folder.
    root = tk.Tk()
    root.withdraw()  # Hide the main window.

    # Ask the user if they want to process an entire folder of PDFs.
    if messagebox.askyesno("Select Processing Mode", "Would you like to process an entire folder of PDFs? Click Yes for folder, No for a single file."):
        # Optionally, if you need to specify the Poppler path (e.g., on Windows), set it here:
        poppler_path = None
        # Example for Windows (uncomment and set your actual path):
        # poppler_path = r"C:\path\to\poppler\bin"
        process_folder(poppler_path)
    else:
        file_path = filedialog.askopenfilename(title="Select a PDF file",
                                               filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            messagebox.showinfo("Info", "No file selected.")
            return

        # Optionally, if you need to specify the Poppler path (e.g., on Windows), set it here:
        poppler_path = None
        # Example for Windows (uncomment and set your actual path):
        # poppler_path = r"C:\path\to\poppler\bin"
        process_pdf(file_path, poppler_path)

    root.destroy()


if __name__ == "__main__":
    main()
