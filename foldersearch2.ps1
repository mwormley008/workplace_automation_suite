
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

    if ($doc.Extension -eq '.doc') {
        # If the file is a .doc, convert it to .docx
        Write-Host "Converting $($doc.Name) to .docx"
        $newDocPath = ConvertTo-Docx -docPath $doc.FullName

        # Open the converted .docx file
        Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$newDocPath`""
        Write-Host "Opening $($doc.Name) as .docx"
    } else {
        # If the file is already a .docx, open it as usual
        Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$($doc.FullName)`""
        Write-Host "Opening $($doc.Name)"
    }
}
elseif ($docFile.Count -gt 1) 
{
    # ... (your existing logic for when there's more than one file)

    foreach ($selected in $selectedFiles)
    {
        $doc = $docFile[([int]$selected) - 1]

        if ($doc.Extension -eq '.doc') {
            # If the file is a .doc, convert it to .docx
            Write-Host "Converting $($doc.Name) to .docx"
            $newDocPath = ConvertTo-Docx -docPath $doc.FullName

            # Open the converted .docx file
            Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$newDocPath`""
            Write-Host "Opening $($doc.Name) as .docx"
        } else {
            # If the file is already a .docx, open it as usual
            Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$($doc.FullName)`""
            Write-Host "Opening $($doc.Name)"
        }
    }
}
else
{
    Write-Host "No file with name `"$fileName`" found."
}
# At the end of your script
Read-Host -Prompt "Press Enter to exit"
