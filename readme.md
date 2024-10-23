# Cladogram Matrix Auto-correction for BIO1130 Lab3 e-matrix

This Python script automates the correction of e-matrices submitted by students of the BIO1130 lab3. It processes each student's submitted matrix, compares it against a correction matrix, and exports the corrected version as a PDF. This script has been tested on macOS.

## Prerequisites

Before using this script, please ensure the following:

1. **System Compatibility**: 
   - The script has been tested on macOS. Compatibility with Windows systems has not been verified.

2. **Password Removal**:
   - The correction matrix (`Correction matrix EN Prot - 2024.xlsx`) must have its password removed.
   - To remove the password, open the file in Excel, select "Save As," click on "Options," remove the password, and save it as a new file.

3. **Adjust Privacy Settings on macOS**:
   - Grant Full Disk Access to the Terminal app:
     - Go to **System Settings** > **Privacy & Security**.
     - Under **Privacy**, find **Full Disk Access** and add the **Terminal** app.

4. **Python Installation**:
   - Ensure Python 3 is installed on your system:
     ```
     python3 --version
     ```
   - If Python is not installed, you can install it using Homebrew:
     ```
     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
     brew install python
     ```

## Getting Started

### Step 1: Download the Script

Download the script named `cladogram_matrix_autocorrection.py` from this repository and save it to a location on your desktop.

### Step 2: Prepare the Input Folder

1. Download all students' matrix files into a folder named `input folder` on your desktop.

2. Adjust the file paths in the command below to match your environment:

   - Replace `/Users/ruizhang/Desktop/input folder` with the full path to your input folder.
   - Specify the file path for the password-removed correction matrix.
   - Specify the folder where you want the auto-corrected matrices to be saved.

### Step 3: Run the Script

Update the paths in the command below and run it in the Terminal:

```bash
source_folder="/Users/ruizhang/Desktop/input folder"
excel_template="/Users/ruizhang/Desktop/Correction matrix EN Prot - 2024 password removed.xlsx"
target_folder="/Users/ruizhang/Desktop/output folder"

python3 cladogram_matrix_autocorrection.py "$source_folder" "$excel_template" "$target_folder"
```

### What the Script Does

1. Iterates through each Excel file in the specified `input folder`.
2. Uses the password-removed correction matrix as a template for comparing and correcting each student's matrix.
3. Extracts the content of cell `H1` from each student's file to name the output PDF.
4. Saves the corrected matrix in the specified `output folder` with a filename in the format `XX_corrected.pdf`, where `XX` is the value from cell `H1`.


## Toubleshooting
1. The program might ask you to "grant access" during each iteration of the student's file, you can adjust the system privacy seeting to allow full disk access by Excel to solve this issue.
2. if the student didn't not fill in their student number in their excel file, you program will prompt `Skipping file example_file.xlsx: H1 is empty.`In this case, you might want to manually fill in the student number for the student and run the program again.


## Example Command

```bash
source_folder="/Users/ruizhang/Desktop/input folder"
excel_template="/Users/ruizhang/Desktop/Correction matrix EN Prot - 2024 password removed.xlsx"
target_folder="/Users/ruizhang/Desktop/output folder"

python3 cladogram_matrix_autocorrection.py "$source_folder" "$excel_template" "$target_folder"
```

This will process all files in the specified `input folder` and save the corrected PDFs in the `output folder`.
