import os
import comtypes.client
import PyPDF2
import re
import tkinter as tk
from tkinter import filedialog, Tk, BooleanVar, Checkbutton, Button

# 1. Function to check and install missing dependencies
def install_dependencies():
    required_libraries = ['comtypes', 'PyPDF2', 're']
    for lib in required_libraries:
        try:
            __import__(lib)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

install_dependencies()

# 2. Create folder selection dialog box for input and output directories
def select_directory(title="Select Folder"):
    Tk().withdraw()  # Hide the root window
    folder_selected = filedialog.askdirectory(title=title)
    return folder_selected

# 3. Convert the documents to PDF
def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        # Ensure the file exists before trying to convert
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"The file {docx_path} does not exist.")

        # Ensure the paths are absolute and properly formatted
        docx_path = os.path.abspath(docx_path)
        pdf_path = os.path.abspath(pdf_path)

        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is for pdf format
        doc.Close()
        word.Quit()
        print(f"Converted: {docx_path} -> {pdf_path}")
    except Exception as e:
        print(f"Failed to convert {docx_path}: {str(e)}")

# 4. Sort Files Alphabetically and Merge PDFs
def merge_pdfs_in_folder(folder_path, output_folder):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    pdf_files.sort()

    if not pdf_files:
        print("No PDF files to merge.")
        return

    merger = PyPDF2.PdfMerger()

    # 5. Get the first PDF's filename and extract the part inside brackets
    first_pdf = pdf_files[0]
    match = re.search(r'\((.*?)\)', first_pdf)
    if match:
        name_part = match.group(1)
    else:
        name_part = "Merged_File"  # Fallback if no brackets found

    # 6. Create the new combined PDF file name with extracted name part
    combined_pdf_filename = f"{name_part}.pdf"
    combined_pdf_path = os.path.join(output_folder, combined_pdf_filename)

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        with open(pdf_path, 'rb') as f:
            merger.append(f)

    with open(combined_pdf_path, 'wb') as combined_pdf:
        merger.write(combined_pdf)

    print(f"Created combined PDF: {combined_pdf_path}")

# 7. Convert and Merge PDFs, while maintaining the directory structure
def convert_and_merge(input_folder, output_folder, delete_word_docs, combine_pdfs):
    for root, dirs, files in os.walk(input_folder):
        relative_path = os.path.relpath(root, input_folder)  # Get the relative path from the input folder
        current_output_folder = os.path.join(output_folder, relative_path)  # Mirror the structure in the output folder

        # Create the corresponding output folder if it doesn't exist
        if not os.path.exists(current_output_folder):
            os.makedirs(current_output_folder)

        pdf_files = []
        for filename in files:
            if filename.endswith(".docx"):
                docx_file = os.path.join(root, filename)
                pdf_file = os.path.join(current_output_folder, filename.replace(".docx", ".pdf"))
                try:
                    convert_docx_to_pdf(docx_file, pdf_file)
                    if delete_word_docs:
                        os.remove(docx_file)  # Delete the original Word document if the option is enabled
                        print(f"Deleted: {docx_file}")
                    pdf_files.append(pdf_file)
                except Exception as e:
                    print(f"Failed to convert {docx_file}: {str(e)}")

        if pdf_files and combine_pdfs:  # If there are PDFs to merge and combining is enabled
            try:
                merge_pdfs_in_folder(current_output_folder, current_output_folder)
            except Exception as e:
                print(f"Failed to merge PDFs in {current_output_folder}: {str(e)}")

# 8. Create GUI to choose options
def open_options_gui():
    root = Tk()
    root.title("Options")

    delete_word_docs_var = BooleanVar(value=True)
    combine_pdfs_var = BooleanVar(value=True)

    Checkbutton(root, text="Delete Word Documents after conversion", variable=delete_word_docs_var).pack(anchor='w')
    Checkbutton(root, text="Combine PDFs after conversion", variable=combine_pdfs_var).pack(anchor='w')

    def on_submit():
        root.quit()  # Close the window
        root.destroy()  # Destroy the window object
        start_process(delete_word_docs_var.get(), combine_pdfs_var.get())

    Button(root, text="Submit", command=on_submit).pack()
    root.mainloop()

# 9. Main Program: Select Input and Output Directories and Process
def start_process(delete_word_docs, combine_pdfs):
    # Select input directory
    input_folder = select_directory("Select the folder containing DOCX files")
    if not input_folder:
        print("No input folder selected.")
        exit()

    # Select output directory
    output_folder = select_directory("Select the output folder for merged PDF")
    if not output_folder:
        print("No output folder selected.")
        exit()

    # Perform conversion and merging with options
    convert_and_merge(input_folder, output_folder, delete_word_docs, combine_pdfs)

if __name__ == "__main__":
    open_options_gui()
