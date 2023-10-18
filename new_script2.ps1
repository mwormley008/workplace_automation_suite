$fileName = Read-Host -Prompt 'Enter a File Name'

# Set the path to the folder you want to search
$folderPath = '\\WBR\shared\G702 & G703 Forms'

# Search for both .xls and .xlsx files that match the user's input in the specified folder
$excelFile = Get-ChildItem -Path $folderPath -File | Where-Object { ($_.Name -like "*$fileName*.xls") -or ($_.Name -like "*$fileName*.xlsx") }

# Check if any files were found
if ($excelFile.Count -ge 1) 
{
    & {
        for ($i = 0; $i -lt $excelFile.Count; $i++)
        {
            [pscustomobject]@{
                Index = $i + 1
                FileName = $excelFile[$i].Name  # Changed from .BaseName to .Name
                Created = $excelFile[$i].CreationTime
            }
        }
    } | Format-Table -AutoSize | Out-String -Stream

    # Let the user choose which file(s) to open based on the Index
    $selection = Read-Host -Prompt 'Enter the number(s) of the file(s) you want to open'

    # If the user's input contains a comma, split by comma and trim spaces
    # Otherwise, treat it as a single integer
    if ($selection -contains ',') {
        $selectedFiles = $selection -split ',' | ForEach-Object { $_.Trim() }
    } else {
        $selectedFiles = @($selection)
    }

    foreach ($selected in $selectedFiles)
    {
        $fileToOpen = $excelFile[([int]$selected) - 1]
        # Open the selected file with Microsoft Excel
        Start-Process -FilePath "EXCEL.EXE" -ArgumentList "`"$($fileToOpen.FullName)`""
        Write-Host "Opening $($fileToOpen.Name)"
    }
}
else
{
    Write-Host -Object "No file with name `"$fileName`" found."
}
