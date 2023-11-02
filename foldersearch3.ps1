# Define the path to the Python revision script
$revisionScriptPath = "C:\Users\Michael\Desktop\python-work\word_revision.py"

$fileName = Read-Host -Prompt 'Enter a File Name'


# Set the path to the folder you want to search
$folderPath = '\\WBR\shared\Proposals'

# Search for both .doc and .docx files that match the user's input in the specified folder
$docFile = Get-ChildItem -Path $folderPath -File | Where-Object { ($_.Name -like "*$fileName*.doc") -or ($_.Name -like "*$fileName*.docx") }

# Function to convert a .doc file to .docx
function ConvertTo-Docx($docPath) {
    # Create a new Word application
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false  # Set to $true if you want to see Word

    try {
        # Open the .doc file
        $document = $wordApp.Documents.Open($docPath)

        # Set the new file path by changing the extension to .docx
        $newDocPath = [System.IO.Path]::ChangeExtension($docPath, ".docx")

        # Save this file with the .docx extension
        $document.SaveAs([ref] $newDocPath, [ref] 16) # 16 = wdFormatDocumentDefault (i.e., .docx)
        $document.Close()

        # Return the new path
        return $newDocPath
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        # Quit Word (important to prevent memory leaks)
        $wordApp.Quit()
    }
}

# Check if any files were found
if ($docFile.Count -eq 1) 
{
    $doc = $docFile  # There's only one document, so we take that.

    # Path variable for the file to be opened
    Write-Host "One file found: $($doc.Name)"  # Display the file name
    
    $pathToOpen = $doc.FullName

    # Ask if the user wants to run the revision script
    $runRevision = Read-Host -Prompt 'Do you want to run the revision script? (y/n)'
    if ($doc.Extension -eq '.doc') {
        # If the file is a .doc, convert it to .docx
        Write-Host "Converting $($doc.Name) to .docx"
        $newDocPath = ConvertTo-Docx -docPath $doc.FullName
        $pathToOpen = $newDocPath  # We update the path to be opened with the new .docx file path
    }


    # Check if the user agreed to run the revision script
    if ($runRevision -eq 'y') {
        if (Test-Path $revisionScriptPath) {
            Write-Host "Running revision script..."
            # Run the Python script and capture the output, which should be the path to the revised file
            $revisionOutput = python $revisionScriptPath $pathToOpen | Out-String
            $revisedFilePath = $revisionOutput -match "'.*?'" | Out-Null
            $revisedFilePath = $matches[0] -replace "'", ""  # Remove the surrounding single quotes

            # Trim any whitespace and remove any hidden special characters that might be included
            $revisedFilePath = $revisedFilePath.Trim()

            # Add this line for debugging:
            Write-Host "Debug: Revised file path is '$revisedFilePath'"

            if (Test-Path $revisedFilePath) {
                $pathToOpen = $revisedFilePath  # We update the path to be opened with the revised file path
                Write-Host "Document revised and saved as '$revisedFilePath'"

            } else {
                Write-Host "Revised file not found at path: $revisedFilePath. Opening original document."
                # If the revised file is not found, it continues to open the original document.
            }
        } else {
            Write-Host "Revision script path not found. Please check the path: $revisionScriptPath"
            # The script proceeds to open the original document if the revision script is not found.
        }
    }

    # Open the document (original or revised)
    Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$pathToOpen`""
    Write-Host "Opening $($pathToOpen.Name)"

} elseif ($docFile.Count -gt 1) 
{
    # If multiple files match, prompt the user to choose one.
    Write-Host "Multiple files found:"
    $index = 1
    foreach ($file in $docFile) {
        Write-Host "$index. $($file.Name)"
        $index++
    }

    $selectedFileIndex = Read-Host -Prompt 'Enter the number of the file you want to open'
    $selectedFile = $docFile[[int]$selectedFileIndex - 1]

    # Path variable for the file to be opened
    $pathToOpen = $selectedFile.FullName

    # Ask if the user wants to run the revision script
    $runRevision = Read-Host -Prompt 'Do you want to run the revision script? (y/n)'
    if ($selectedFile.Extension -eq '.doc') {
        # If the file is a .doc, convert it to .docx
        Write-Host "Converting $($selectedFile.Name) to .docx"
        $newDocPath = ConvertTo-Docx -docPath $selectedFile.FullName
        $pathToOpen = $newDocPath  # Update the path to the new .docx file
    }

    # Check if the user agreed to run the revision script
    if ($runRevision -eq 'y') {
        if (Test-Path $revisionScriptPath) {
            Write-Host "Running revision script..."
            # Run the Python script and capture the output, which should be the path to the revised file
            $revisionOutput = python $revisionScriptPath $pathToOpen | Out-String
            $revisedFilePath = $revisionOutput -match "'.*?'" | Out-Null
            $revisedFilePath = $matches[0] -replace "'", ""  # Remove the surrounding single quotes

            # Trim any whitespace and remove any hidden special characters that might be included
            $revisedFilePath = $revisedFilePath.Trim()

            if (Test-Path $revisedFilePath) {
                $pathToOpen = $revisedFilePath  # Update the path to be opened with the revised file path
                Write-Host "Document revised and saved as '$revisedFilePath'"
            } else {
                Write-Host "Revised file not found at path: $revisedFilePath. Opening original document."
            }
        } else {
            Write-Host "Revision script path not found. Please check the path: $revisionScriptPath"
        }
    }

    # Open the document (original or revised)
    Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$pathToOpen`""
    Write-Host "Opening $($pathToOpen.Name)"
}
else
{
    Write-Host "No file with name `"$fileName`" found."
}

# At the end of your script
# Read-Host -Prompt "Press Enter to exit"