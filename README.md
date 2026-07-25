# Word to PDF Folder Converter and Merger

A Windows Python utility that converts DOCX files to PDF, mirrors subfolders, optionally merges the PDFs in each folder and can remove the original Word documents.

This is an earlier, more specialised converter. It is useful for controlled folder based conversion jobs where per folder merging is required. For routine batch conversion where retaining source documents matters, the newer ConvertalldocxtoPDF project is the safer starting point.

> [!CAUTION]
> The Delete Word Documents option is selected by default. In the current code, a source DOCX file can be deleted even when its PDF conversion has failed because the conversion function handles the error internally without reporting failure to the deletion step. Leave deletion unticked unless you have a verified backup and have tested the complete workflow on copied files.

## What it does

* Prompts for an input folder and a separate output folder.
* Recursively finds DOCX files.
* Uses the locally installed Microsoft Word application to create PDFs.
* Mirrors the input directory structure in the output location.
* Optionally combines PDFs within each output subfolder.
* Names a combined PDF from text inside parentheses in the first PDF filename, or uses Merged_File.pdf when no match exists.
* Optionally deletes each original DOCX after the conversion attempt.

## Requirements

* Windows
* Microsoft Word desktop installed and activated
* Python 3 with Tkinter
* comtypes
* PyPDF2

Install the dependencies before running:

    python -m pip install comtypes PyPDF2

Preinstallation is important because the current automatic dependency installation path refers to modules that are not imported at that point in the script.

## Run

    python "v4.1 Convert Word to PDF - With Browse and Options Delete and Combine.py"

For a safer first run:

1. Clear Delete Word Documents after conversion.
2. Leave Combine PDFs selected only if a combined file is required.
3. Use copies of the source documents.
4. Choose an empty output folder.
5. Review every converted PDF before changing or removing the originals.

## Merge behaviour

PDFs are sorted alphabetically within each output folder before merging. If the first filename contains parentheses, the text inside the first pair becomes the merged filename. Otherwise, the output is Merged_File.pdf.

This naming rule is specific to the current code and may overwrite a file with the same name in the selected output folder.

## Known limitations

* Only DOCX files are processed. Legacy DOC files are ignored.
* Conversion requires Windows and Microsoft Word.
* The source deletion workflow is unsafe unless the input is backed up and output is independently verified.
* Word processes are created for individual documents, which can make large jobs slow.
* Failed conversions are printed to the console rather than written to a persistent log.
* Merge ordering is alphabetical rather than natural numeric order.
* Existing output and merged files can be overwritten without a separate confirmation.
* The graphical interface may appear unresponsive while conversion is running.

## Licence

Apache License version 2.0. See LICENSE for the complete terms.
