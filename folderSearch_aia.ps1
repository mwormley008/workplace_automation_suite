$fileName = Read-Host -Prompt 'Enter a File Name'

# Set the path to the folder you want to search
$folderPath = '\\WBR\shared\G702 & G703 Forms'

# Search for both .xls and .xlsx files that match the user's input in the specified folder
$excelFile = Get-ChildItem -Path $folderPath -File | Where-Object { ($_.Name -like "*$fileName*.xls") -or ($_.Name -like "*$fileName*.xlsx") }

# Check if any files were found
if ($excelFile) 
{
    # Display all files, even if there's just one
    $fileList = & {
        for ($i = 0; $i -lt $excelFile.Count; $i++)
        {
            [pscustomobject]@{
                Index = $i + 1
                FileName = $excelFile[$i].BaseName
                Created = $excelFile[$i].CreationTime
            }
        }
    } | Format-Table -AutoSize | Out-String -Stream

    # Display the file list
    $fileList | ForEach-Object { Write-Host $_ }

    # Let the user choose which file(s) to open based on the Index
    $selection = Read-Host -Prompt 'Enter file number(s) to open (or leave empty to cancel)'

    # Check if the user input is empty (indicating they want to cancel)
    if ([string]::IsNullOrWhiteSpace($selection)) {
        Write-Host "No selection made. Exiting..."
        return
    }

    # If the user's input contains a comma, split by comma and trim spaces
    # Otherwise, treat it as a single integer
    if ($selection -contains ',') {
        $selectedFiles = $selection -split ',' | ForEach-Object { $_.Trim() }
    } else {
        $selectedFiles = @($selection)
    }

    foreach ($selected in $selectedFiles)
    {
        # Validate the selected index before trying to access the file
        if ($selected -gt 0 -and $selected -le $excelFile.Count) {
            $file = $excelFile[([int]$selected) - 1]
            # Open the selected file with Microsoft Excel
            Start-Process -FilePath "EXCEL.EXE" -ArgumentList "`"$($file.FullName)`""
            Write-Host "Opening $($file.Name)"
        } else {
            Write-Host "Invalid selection: $selected. Please select a number from the list."
        }
    }
}
else
{
    Write-Host -Object "No file with name `"$fileName`" found."
}
