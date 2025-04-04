import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import logging
import subprocess
import sys

# Configure logging
logging.basicConfig(filename='pdf_blank_page_remover2.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Global flag for stopping the process safely
stop_requested = False

def install_dependencies():
    """Installs missing dependencies if the user agrees."""
    try:
        import fitz
        import pdf2image
        import pytesseract
        from PIL import Image
    except ImportError:
        if messagebox.askyesno("Dependencies Missing", "Some dependencies are missing. Do you want to install them?"):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF", "pdf2image", "pytesseract", "Pillow"])
                messagebox.showinfo("Installation Complete", "Dependencies installed. Please restart the application.")
                sys.exit()
            except subprocess.CalledProcessError:
                messagebox.showerror("Installation Failed", "Failed to install dependencies. Please install them manually.")
                sys.exit()

install_dependencies()

def check_blank_page(page, use_ocr=False):
    """Checks if a PDF page is blank."""
    if use_ocr:
        try:
            # Convert the page to an image using a higher DPI for better OCR accuracy.
            images = convert_from_path(page.parent.name, first_page=page.number + 1, last_page=page.number + 1, dpi=200)
            if not images:
                return True  # Consider it blank if no image is generated.
            text = pytesseract.image_to_string(images[0], lang='eng')
            return len(text.strip()) < 5
        except Exception as e:
            logging.error(f"OCR error: {e}")
            return False  # If OCR fails, do not mark as blank.
    else:
        text = page.get_text("text").strip()
        return not text

def update_progress(progress_var, progress_label, new_value, filename):
    """Updates the progress variable and label in the main thread."""
    progress_var.set(new_value)
    progress_label.config(text=f"Processed: {filename}")

def process_pdf(filepath, progress_var, progress_label, log_file):
    """Processes a single PDF file, removing blank pages."""
    try:
        doc = fitz.open(filepath)
        total_pages = len(doc)
        blank_pages = []

        # First, check for blank pages using text extraction.
        for page_num in range(total_pages):
            page = doc[page_num]
            if check_blank_page(page, use_ocr=False):
                blank_pages.append(page_num)

        # Refine blank page detection using OCR.
        blank_pages_ocr = []
        for page_num in blank_pages:
            page = doc[page_num]
            if check_blank_page(page, use_ocr=True):
                blank_pages_ocr.append(page_num)
        blank_pages = blank_pages_ocr  # Only keep pages confirmed blank by OCR

        if blank_pages:
            temp_filepath = filepath + ".bak"
            shutil.copyfile(filepath, temp_filepath)  # Create a backup
            new_doc = fitz.open()
            for page_num in range(total_pages):
                if page_num not in blank_pages:
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            new_doc.save(filepath)
            new_doc.close()
            doc.close()
            os.remove(temp_filepath)  # Remove backup after successful save
            log_file.write(f"File: {os.path.basename(filepath)}, Total pages: {total_pages}, Removed pages: {len(blank_pages)}, Status: Modified\n")
            logging.info(f"File: {os.path.basename(filepath)}, Total pages: {total_pages}, Removed pages: {len(blank_pages)}, Status: Modified")
        else:
            doc.close()
            log_file.write(f"File: {os.path.basename(filepath)}, Total pages: {total_pages}, Removed pages: 0, Status: No changes\n")
            logging.info(f"File: {os.path.basename(filepath)}, Total pages: {total_pages}, Removed pages: 0, Status: No changes")
    except Exception as e:
        log_file.write(f"Error processing {os.path.basename(filepath)}: {e}\n")
        logging.error(f"Error processing {os.path.basename(filepath)}: {e}")
    finally:
        # Schedule the progress update to run in the main thread.
        new_value = progress_var.get() + 1
        root.after(0, update_progress, progress_var, progress_label, new_value, os.path.basename(filepath))

def process_files(folder_path, progress_window, progress_var, progress_label, progress_bar):
    """Processes all PDF files in the selected folder."""
    global stop_requested
    files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    if not files:
        root.after(0, progress_window.destroy)
        root.after(0, messagebox.showinfo, "No PDF Files", "No PDF files found in the selected folder.")
        return

    # Set the maximum value of the progress bar.
    progress_bar.config(maximum=len(files))
    progress_var.set(0)
    log_file = open("pdf_processing2_log.txt", "w")

    for filepath in files:
        if stop_requested:
            break  # Stop processing if the stop flag is set.
        process_pdf(filepath, progress_var, progress_label, log_file)

    log_file.close()
    # Determine finish message based on whether processing was stopped.
    finish_message = "PDF processing stopped by user." if stop_requested else "PDF processing finished."
    # Schedule closing of the progress window and display the finish message.
    root.after(0, progress_window.destroy)
    root.after(0, messagebox.showinfo, "Processing Complete", finish_message)

def stop_processing(stop_button):
    """Sets the global flag to stop processing and disables the stop button."""
    global stop_requested
    stop_requested = True
    stop_button.config(state="disabled")

def select_folder():
    """Opens a folder selection dialog and starts processing."""
    folder_path = filedialog.askdirectory(title="Select Folder Containing PDFs")
    if folder_path:
        # Create a new progress window.
        progress_window = tk.Toplevel(root)
        progress_window.title("Processing PDFs...")
        progress_window.geometry("350x200")
        progress_window.resizable(True, True)

        # --- Creating the progress bar and label in the progress window ---
        # Create an IntVar for tracking progress
        progress_var = tk.IntVar(value=0)

        # Create the progress bar widget, linking it to the IntVar
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", mode="determinate",
                                       variable=progress_var)
        progress_bar.pack(pady=10, fill="x", padx=10)

        # Create a label to show which file is being processed
        # progress_label = tk.Label(progress_window, text="Processing...")
        # progress_label.pack(pady=5)

        # Create a label to show which file is being processed
        progress_label = tk.Label(progress_window, text="Processing...")
        progress_label.pack(pady=5)
        progress_window.lift()  # Bring the progress window to the front
        progress_window.attributes('-topmost', True)
        progress_window.update_idletasks()

        # Create the Stop button.
        stop_button = tk.Button(progress_window, text="Stop", command=lambda: stop_processing(stop_button))
        stop_button.pack(pady=5)

        # Start processing files in a separate thread.
        threading.Thread(target=process_files, args=(folder_path, progress_window, progress_var, progress_label, progress_bar), daemon=True).start()

# Set up the main window.
root = tk.Tk()
root.title("PDF Blank Page Remover")

# Set a custom icon if available.
try:
    root.iconbitmap("icon.ico")
except:
    pass

select_button = tk.Button(root, text="Select Folder", command=select_folder, padx=20, pady=10)
select_button.pack(pady=20)

root.mainloop()
