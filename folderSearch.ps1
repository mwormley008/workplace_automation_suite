$fileName = Read-Host -Prompt 'Enter a File Name'

# Set the path to the folder you want to search
$folderPath = '\\WBR\shared\Proposals'

# Search for both .doc and .docx files that match the user's input in the specified folder
$docFile = Get-ChildItem -Path $folderPath -File | Where-Object { ($_.Name -like "*$fileName*.doc") -or ($_.Name -like "*$fileName*.docx") }

# Check if any files were found
if ($docFile.Count -eq 1) 
{
    # If only one file is found, open it immediately
    Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$($docFile.FullName)`""
    Write-Host "Opening $($docFile.Name)"
}
elseif ($docFile.Count -gt 1) 
{
    & {
        for ($i = 0; $i -lt $docFile.Count; $i++)
        {
            [pscustomobject]@{
                Index = $i + 1
                DocName = $docFile[$i].BaseName
                Created = $docFile[$i].CreationTime
            }
        }
    } | Format-Table -AutoSize | Out-String -Stream

    # Let the user choose which file(s) to open based on the Index
    $selection = Read-Host -Prompt 'Enter DOC number(s) to open'

    # If the user's input contains a comma, split by comma and trim spaces
    # Otherwise, treat it as a single integer
    if ($selection -contains ',') {
        $selectedFiles = $selection -split ',' | ForEach-Object { $_.Trim() }
    } else {
        $selectedFiles = @($selection)
    }

    foreach ($selected in $selectedFiles)
    {
        $doc = $docFile[([int]$selected) - 1]
        # Open the selected file with Microsoft Word
        Start-Process -FilePath "WINWORD.EXE" -ArgumentList "`"$($doc.FullName)`""
        Write-Host "Opening $($doc.Name)"
    }
}
else
{
    Write-Host -Object "No file with name `"$fileName`" found."
}
