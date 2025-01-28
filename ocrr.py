import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import os
from tkinter import (
    Tk,
    filedialog,
    messagebox,
    Label,
    Entry,
    Button,
    StringVar,
    Text,
    Scrollbar,
    ttk,
    Menu,
    Checkbutton,
    BooleanVar,
)
from tkinter import font as tkFont
import threading

# Global variables for progress
progress = 0
total_files = 0

# Function to extract text from a PDF file using OCR
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text += pytesseract.image_to_string(img)
        pdf_document.close()
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
    return text

# Function to search for target words in PDF files
def find_word_in_pdfs(directory, target_words, case_sensitive):
    global progress, total_files
    matching_files = []
    target_words = [word.strip() for word in target_words.split(",")]
    files = [f for f in os.listdir(directory) if f.endswith(".pdf")]
    total_files = len(files)
    progress = 0

    for filename in files:
        pdf_path = os.path.join(directory, filename)
        text = extract_text_from_pdf(pdf_path)
        for target_word in target_words:
            if case_sensitive:
                if target_word in text:
                    matching_files.append(pdf_path)
                    break
            else:
                if target_word.lower() in text.lower():
                    matching_files.append(pdf_path)
                    break
        progress += 1
        update_progress()

    return matching_files

# Function to update the progress bar and percentage
def update_progress():
    progress_value = (progress / total_files) * 100 if total_files > 0 else 0
    progress_bar["value"] = progress_value
    progress_label.config(text=f"Progress: {int(progress_value)}%")
    if progress_value >= 100:
        progress_label.config(text="Search Complete!")

# Function to select a folder
def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_path.set(folder)

# Function to start the search process
def start_search():
    target_words = word_entry.get().strip()
    folder = folder_path.get()
    case_sensitive = case_sensitive_var.get()

    if not target_words:
        messagebox.showwarning("Input Error", "Please enter target words.")
        return
    if not folder:
        messagebox.showwarning("Input Error", "Please select a folder.")
        return

    # Clear previous results
    results_text.delete(1.0, "end")
    progress_bar["value"] = 0
    progress_label.config(text="Progress: 0%")

    # Start search in a separate thread
    search_thread = threading.Thread(
        target=lambda: display_results(find_word_in_pdfs(folder, target_words, case_sensitive))
    )
    search_thread.start()

# Function to display the results
def display_results(matching_files):
    if matching_files:
        result_text = "PDFs containing the target words:\n"
        result_text += "\n".join(matching_files)
    else:
        result_text = "No PDFs found containing the target words."
    results_text.insert("end", result_text)

# Function to clear all fields
def clear_all():
    word_entry.set("")
    folder_path.set("")
    results_text.delete(1.0, "end")
    progress_bar["value"] = 0
    progress_label.config(text="Progress: 0%")

# Function to save results to a file
def save_results():
    result_text = results_text.get(1.0, "end")
    if result_text.strip():
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, "w") as file:
                file.write(result_text)
            messagebox.showinfo("Save Results", "Results saved successfully.")
    else:
        messagebox.showwarning("Save Results", "No results to save.")

# Function to exit the application
def exit_app():
    root.quit()

# Create the main application window
root = Tk()
root.title("PDF Word Search")
root.geometry("800x600")
root.configure(bg="#f0f8ff")  # Light blue background

# Variables to store user input
folder_path = StringVar()
word_entry = StringVar()
case_sensitive_var = BooleanVar()

# GUI Elements
Label(
    root, text="Enter the word(s) to search for (comma-separated):", bg="#f0f8ff"
).grid(row=0, column=0, padx=10, pady=10, sticky="w")
Entry(root, textvariable=word_entry, width=60, font=("Arial", 12)).grid(
    row=0, column=1, padx=10, pady=10, sticky="w"
)

Label(root, text="Select folder containing PDFs:", bg="#f0f8ff").grid(
    row=1, column=0, padx=10, pady=10, sticky="w"
)
Entry(root, textvariable=folder_path, width=60, font=("Arial", 12)).grid(
    row=1, column=1, padx=10, pady=10, sticky="w"
)
Button(
    root, text="Browse", command=select_folder, bg="#4682b4", fg="white", font=("Arial", 10)
).grid(row=1, column=2, padx=10, pady=10)

Checkbutton(
    root, text="Case Sensitive", variable=case_sensitive_var, bg="#f0f8ff"
).grid(row=2, column=0, padx=10, pady=10, sticky="w")
Button(
    root, text="Search", command=start_search, bg="#32cd32", fg="white", font=("Arial", 10)
).grid(row=2, column=1, padx=10, pady=10, sticky="w")
Button(
    root, text="Clear All", command=clear_all, bg="#ff8c00", fg="white", font=("Arial", 10)
).grid(row=2, column=2, padx=10, pady=10, sticky="w")
Button(
    root, text="Exit", command=exit_app, bg="#dc143c", fg="white", font=("Arial", 10)
).grid(row=2, column=3, padx=10, pady=10, sticky="w")

# Results display area
results_text = Text(root, wrap="word", height=15, width=90, font=("Arial", 10))
results_text.grid(row=3, column=0, columnspan=4, padx=10, pady=10)

# Scrollbar for results
scrollbar = Scrollbar(root, command=results_text.yview)
scrollbar.grid(row=3, column=4, sticky="ns")
results_text.config(yscrollcommand=scrollbar.set)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=700, mode="determinate")
progress_bar.grid(row=4, column=0, columnspan=4, padx=10, pady=10)

# Progress percentage label
progress_label = Label(root, text="Progress: 0%", bg="#f0f8ff")
progress_label.grid(row=5, column=0, columnspan=4, padx=10, pady=10)

# Run the application
root.mainloop()
