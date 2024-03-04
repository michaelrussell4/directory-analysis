## Directory Analysis Script (PowerShell)

**Description:**

This script analyzes a specified directory and its subfolders, generating a CSV file with detailed information about each file. It also creates a log file to track any errors encountered during the analysis.

**Features:**

* Analyzes files and folders recursively.
* Extracts information like filename, size, extension, and SHA256 hash.
* Identifies embedded files within Microsoft Office documents (DOCX, DOCX, PPTX, PPT).
* Handles ZIP archives.
* Logs errors and timestamps.

**Requirements:**

* PowerShell 5.1 or later
* System.IO.Compression.FileSystem assembly (usually included by default)

**Instructions:**

1. Replace `test_cases` with the actual directory path you want to analyze.
2. Save the script as a `.ps1` file (e.g., `analyze_directory.ps1`).
3. Run the script from the PowerShell console: `powershell -ExecutionPolicy Bypass -File analyze_directory.ps1`

**Output:**

* Two files will be created in the same directory as the script:
    * `File_Inventory_for_<dir_name>.csv`: Contains the detailed file information in CSV format.
    * `Log_for_<dir_name>.txt`: Logs any errors encountered during the analysis.

**Note:** Running external tools like `Word.application` might require administrative privileges depending on your system configuration.
