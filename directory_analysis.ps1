# This script analyzes a directory and its subfolders, generating a CSV file with detailed information about each file.
# It also creates a log file to track any errors encountered during the analysis.

# Load the System.IO.Compression.FileSystem assembly (usually loaded by default)
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Define the character encoding used throughout the script
$encoding = [System.Text.Encoding]::UTF8

# Define the directory to be analyzed (replace with the actual path)
$DIRECTORY = "test_cases"

# Define the environment variable (e.g., "TEST", "PRODUCTION")
$ENVIRONMENT = "TEST"

# Write a message to the console indicating the directory being analyzed
Write-Host "Analyzing $DIRECTORY"

# Extract the directory name from the full path
$dirName = Split-Path -Leaf $DIRECTORY

# Handle cases where the directory path ends with a colon (e.g., drive letter)
if ($DIRECTORY[-1] -eq ":") {
    $dirName = $DIRECTORY[0]
}

# Define the output file name for the CSV data
$outputFileName = "File_Inventory_for_$($dirName).csv"

# Define the log file name for error tracking
$logFileName = "Log_for $($dirName).txt"

# Function to generate a random temporary path
function Get-RandomTempPath {
    # Combine the system's temporary path with a random file name
    return [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
}

# Function to extract embedded file names from Microsoft Office documents (DOCX, DOCX)
function Get-EmbFileNameFromCOM($MSFile) {
    # Get the file extension
    $extension = [System.IO.Path]::GetExtension($MSFile)

    try {
        # Check if the file is a Word document
        if ($extension -eq ".docx" -or $extension -eq ".doc") {
            # Create a new instance of the Word application
            $mw = New-Object -ComObject Word.application
            # Open the document
            $doc = $mw.Documents.Open($MSFile, $false, $true)

            $embeddedFileName = ""

            # Loop through all fields in the document
            foreach ($field in $doc.Fields) {
                # Check if the field type is an embedded object
                if ($field.Type -eq 58) {
                    # Extract the embedded file name from the field
                    $embeddedFileName = $field.OLEFormat.IconLabel
                    break
                }
            }
        }
    } finally {
        # Ensure proper resource disposal even if exceptions occur
        if ($doc) { $doc.Close() }
        if ($mw) { $mw.Quit() }
    }

    return $embeddedFileName
}


#region Function_Get-FileHash
<#
.SYNOPSIS
Calculates the SHA256 hash of a specified file.

.DESCRIPTION
This function opens a file, computes its SHA256 hash using the System.Security.Cryptography namespace,
and returns the hash as a lowercase string without hyphens.

.PARAMETER FilePath
    Specifies the path to the file for which to calculate the hash.
    Type: String
    Mandatory: True

.OUTPUTS
    The SHA256 hash of the file, as a lowercase string without hyphens.

.EXAMPLE
    Get-FileHash -FilePath "C:\Documents\test.txt"

.NOTES
    - Requires PowerShell version 5.1 or later.
    - Error handling is included to handle potential file access or hashing issues.
#>
function Get-FileHash {
    [CmdletBinding()]  # Enforce parameter binding and positional argument handling
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    try {
        # Create a SHA256 hash algorithm object
        $hashAlgorithm = [System.Security.Cryptography.SHA256]::Create()

        # Open the file stream for reading
        $fileStream = [System.IO.File]::OpenRead($FilePath)

        # Compute the hash of the file content
        $hashBytes = $hashAlgorithm.ComputeHash($fileStream)

        # Close the file stream
        $fileStream.Close()

        # Convert the hash bytes to a string representation
        $hashString = [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()

        return $hashString
    }
    catch {
        Write-Error "Failed to compute hash for file: $FilePath"
        return $null
    }
}
#endregion Function_Get-FileHash


# <summary>
# Extracts embedded files from Microsoft Office applications (DOCX, DOCX, PPTX, PPT) found within the specified file.
# </summary>
# <param name="appFileName">Path to the Microsoft Office application file (DOCX, DOCX, PPTX, PPT).</param>
# <param name="parent">Parent directory of the file containing the embedded files.</param>
# <param name="fullPath">Full path to the file containing the embedded files.</param>
# <returns>An array of PSCustomObjects containing information about the extracted embedded files.</returns>
function Get-EmbeddedFilesFromMSApplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appFileName,
        [Parameter(Mandatory = $true)]
        [string]$parent,
        [Parameter(Mandatory = $true)]
        [string]$fullPath
    )

    try {
        # Create a temporary directory for extracted files.
        $temporaryDirectory = Get-RandomTempPath

        # Extract the contents of the application file to the temporary directory.
        [System.IO.Compression.ZipFile]::ExtractToDirectory($appFileName, $temporaryDirectory, $encoding)

        # Get the path to the embedded folder within the temporary directory.
        $embeddedFolder = (Get-ChildItem -Path $temporaryDirectory -Filter "*embeddings*" -Recurse -Directory).FullName

        # Check if the embedded folder exists.
        if (-not ($embeddedFolder -and $embeddedFolder -is [string])) {
            return @()
        }

        # Get all binary files within the embedded folder.
        $embeddedBinFiles = Get-ChildItem -Path $embeddedFolder -File

        # Regular expression for extracting embedded file names.
        $extractFileNamePattern = '\\([\w\.~\{\}-]+\\)+([\w\.-]+)'

        # Array to store information about extracted embedded files.
        $files = @()

        foreach ($embeddedBinFile in $embeddedBinFiles) {
            # Read the content of the embedded binary file.
            $fileText = Get-Content -Path $embeddedBinFile.FullName -Encoding ascii -Raw

            # Initialize variable to store extracted file name.
            $fileName = ""

            # Attempt to extract the file name from the content using the regular expression.
            $fileText | Select-String -Pattern $extractFileNamePattern | ForEach-Object {
                $match = $_.Matches[0]
                $fileName = $match.Groups[2].Value
                return
            }

            # If the file name couldn't be extracted, use the application file name as a fallback.
            if ($fileName -eq "") {
                $fileName = Get-EmbFileNameFromCOM($appFileName)
            }

            # Get the parent file name.
            $parentFile = $appFileName.Split('\')[-1]

            # Construct the full path of the embedded file relative to the original file's parent.
            $parentFileRFP = Join-Path -Path $fullPath -ChildPath $parentFile

            # Attempt to extract classification information from the file name using a regular expression.
            $fileClassification = [regex]::Match($fileName, "\((.*?)\)")

            # Determine the size of the embedded file content.
            if ($null -eq $fileText -or $fileText -eq "") {
                $fileBytes = 0
            }
            else {
                $fileBytes = [System.Text.Encoding]::UTF8.GetBytes($fileText).Length
                # Limit the size to avoid potential memory issues.
                if ($fileBytes -gt 2000) {
                    $fileBytes -= 2000
                }
            }

            # Get the extension of the embedded file.
            $fileExtension = [System.IO.Path]::GetExtension($fileName)

            # Create a PSCustomObject to store information about the extracted embedded file.
            $file = [PSCustomObject]@{
                Environment = $ENVIRONMENT
                Parent = $parent
                Name = $fileName
                FullPath = $parentFileRFP
                Classification = $fileClassification
                Bytes = $fileBytes
                Extension = $fileExtension
                SHA256Hash = $null
                HasEmbedded = $false
                IsEmbedded = $true
            }

            # Add the PSCustomObject to the array of extracted embedded files.
            $files += $file
        }
    }
    catch [System.IO.InvalidDataException] {
        # Handle invalid data exceptions and return an empty
        # Log an error message if the temporary directory cannot be removed.
        if (Test-Path -Path $temporaryDirectory) {
            Remove-Item -Path $temporaryDirectory -Recurse -Force
            Write-Error "Failed to remove temporary directory: $temporaryDirectory"
        }

        # Return the array of PSCustomObjects containing information about the extracted embedded files.
        return @()
    }
    finally {
        # Ensure the temporary directory is removed even if an exception occurs.
        if (Test-Path -Path $temporaryDirectory) {
            Remove-Item -Path $temporaryDirectory -Recurse -Force
        }
    }
    return $files
}

# <summary>
# Copies a directory and its contents to a randomly generated temporary location.
# </summary>
# <param name="path">The path to the directory to copy.</param>
# <returns>The full path to the temporary location where the directory was copied.</returns>
function Copy-DirectoryToTemp {
    param (
        [Parameter(Mandatory = $true)]
        [string]$path
    )

    # Generate a random temporary directory path.
    $temporaryDirectory = Get-RandomTempPath

    # Copy the specified directory and its contents recursively to the temporary location,
    # overwriting any existing files and folders.
    Copy-Item -Path $path -Destination $temporaryDirectory -Recurse -Force

    # Return the full path to the temporary directory.
    return $temporaryDirectory
}


# <summary>
# Recursively scans a folder and its subfolders, extracting information about files and embedded files (if applicable) and exporting it to a CSV file.
# </summary>
# <param name="path">The path to the folder to be scanned.</param>
# <param name="parent">The name of the parent folder. (Optional)</param>
# <param name="fullPath">The full path to the file being processed. (Optional)</param>
function Get-AllFilesInFolder {
    param (
        [Parameter(Mandatory = $true)]
        [string]$path,
        [string]$parent = $null,
        [string]$fullPath = $null
    )

    # If parent folder is not specified, extract it from the path.
    if (-not $parent) {
        $parent = Split-Path -Leaf $path
    }

    # If full path is not specified, set it to the provided path.
    if (-not $fullPath) {
        $fullPath = $path
    }

    # Ensure full path starts with "\\?\" to handle long paths.
    if (!$path.StartsWith("\\?\")) {
        $path = "\\?\$path"
    }

    # Recursively get all files (not folders) within the specified path.
    Get-ChildItem -Force -Path $path -Recurse | Where-Object { $_.PSIsContainer -eq $false } | ForEach-Object {
        # Get the relative path to the file within the current path.
        $fromPathToFile = $_.DirectoryName.Substring($path.Length)
        $fullPathToFile = Join-Path -Path $fullPath -ChildPath $fromPathToFile

        try {
            $fileName = $_.Name

            # Process ZIP files:
            if ($fileName -like "*.zip") {
                $temporaryDirectory = Get-RandomTempPath

                try {
                    # Extract ZIP contents to a temporary directory.
                    [System.IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, $temporaryDirectory, [System.Text.Encoding]::Default)
                } catch [System.IO.PathTooLongException] {
                    Write-Host "Warning: A path within a zipped folder was too long and threw an error. ($fullPathToFile)"
                }

                # Recursively scan extracted files.
                $zippedFolderRFP = Join-Path -Path $fullPathToFile -ChildPath $_.Name
                Get-AllFilesInFolder $temporaryDirectory -parent $parent -fullPath $zippedFolderRFP

                # Clean up the temporary directory.
                if (Test-Path -Path $temporaryDirectory) {
                    Remove-Item -Path $temporaryDirectory -Recurse -Force
                }
            }
            else {
                # Process non-ZIP files:
                $fileItem = $_ | Select-Object Environment, Parent, Name, FullPath, Classification, Bytes, Extension, SHA256Hash, HasEmbedded, IsEmbedded

                # Populate file information.
                $fileItem.Environment = $ENVIRONMENT  # Assuming $ENVIRONMENT is defined elsewhere
                $fileItem.Parent = $parent
                $fileItem.FullPath = $fullPathToFile
                $fileItem.Classification = [regex]::Match($_.Name, "\(([\w\s]*?)\)")
                $fileItem.Bytes = $_.Length
                $fileItem.SHA256Hash = Get-FileHash -FilePath $_.FullName
                $fileItem.HasEmbedded = $false
                $fileItem.IsEmbedded = $false

                # Check for embedded files in MS Office documents:
                $extension = $fileItem.Extension
                if ($extension -eq ".docx" -or $extension -eq ".doc" -or $extension -eq ".pptx" -or $extension -eq ".ppt") {
                    $embeddedFiles = Get-EmbeddedFilesFromMSApplication -appFileName $_.FullName -parent $parent -fullPath $fullPathToFile

                    if ($embeddedFiles) {
                        $fileItem.HasEmbedded = $true
                        $embeddedFilesSizesSum = $embeddedFiles | ForEach-Object { $_.Bytes } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                        $fileItem.Bytes -= [Math]::Min($fileItem.Bytes, $embeddedFilesSizesSum)
                        $embeddedFiles | Export-Csv -Path $outputFileName -NoTypeInformation -Append -Encoding UTF8
                    }
                }

                # Export file information to CSV.
                $fileItem | Export-Csv -Path $outputFileName -NoTypeInformation -Append -Encoding UTF8
            }
        } catch {
            # Throw an exception with additional context for troubleshooting.
            throw @($_.Exception.Message, $fullPathToFile, $fileName)
        }
    }

    # Remove the temporary directory created for ZIP extraction (if applicable).
    if ($path.StartsWith([System.IO.Path]::GetTempPath())) {
        Remove-Item -Path $path -Recurse -Force
    }
}

# <summary>
# Removes any existing output files (log and CSV), ensuring a clean starting point.
# </summary>
function Remove-OutputFiles {
    # Check for and remove the log file.
    if (Test-Path $logFileName) {
        Remove-Item $logFileName
    }

    # Check for and remove the output CSV file.
    if (Test-Path $outputFileName) {
        Remove-Item $outputFileName
    }
}

# <summary>
# Appends a timestamp to the log file, providing a reference for execution events.
# </summary>
function Write-TimestampToOutput {
    # Get the current timestamp in a readable format.
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Append the timestamp to the log file with a descriptive message.
    Write-Output "Timestamp: $timestamp" >> $logFileName
}

# <summary>
# Writes an error message to the log file with details, aiding in troubleshooting.
# </summary>
# <param name="folder">The folder where the error occurred.</param>
# <param name="exception">The exception object containing error details.</param>
function Write-ErrorLog {
    param(
        [string]$folder,
        [Exception]$exception
    )

    # Get the current timestamp for the error log.
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Construct a detailed error message with formatting for clarity.
    $errorMessage = @"
Error occurred in folder: $folder
   @Timestamp: $timestamp

   Exception:
   $($exception.Message)

   Error details:
   $($exception)
"@

    # Write the error message to the log file, redirecting errors as well.
    Write-Error $errorMessage 2>> $logFileName
}

# <summary>
# Displays a progress bar in the console for user feedback.
# </summary>
# <param name="Current">The current item being processed.</param>
# <param name="Total">The total number of items to be processed.</param>
# <param name="DirectoryName">The name of the directory being analyzed.</param>
function Write-ProgressUpdate {
    param (
        [int]$Current,
        [int]$Total,
        [string]$DirectoryName
    )

    # Calculate the percentage completion for the progress bar.
    $PercentComplete = [Math]::Round(($Current / $Total) * 100)

    # Display the progress bar with descriptive text.
    Write-Progress -Activity "Analyzing $DirectoryName" -Status "Progress" -PercentComplete $PercentComplete
}

# Get the file data for each first-level parent in the directory.
$parentFolders = Get-ChildItem -Path $DIRECTORY -Directory
$numberOfFolders = $parentFolders.Length
$i = 0
Remove-OutputFiles
Write-TimestampToOutput
$parentFolders | ForEach-Object {
    try {
        $folder = $_.FullName
        Get-AllFilesInFolder $folder
    } catch {
        Write-ErrorLog -folder $folder -exception $_.Exception
    }
    Write-ProgressUpdate -Current $i -Total $numberOfFolders -DirectoryName $dirName
    $i += 1
}
Write-TimestampToOutput