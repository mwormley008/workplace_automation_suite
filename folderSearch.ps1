$fileName = Read-Host -Prompt 'Enter a File Name'

# Set the path to the folder you want to search
$folderPath = '\\WBR\shared\Proposals'

# Search for .doc files that match the user's input in the specified folder
if ($docFile = Get-ChildItem -Path $folderPath -Filter "*$fileName*.doc" -File)
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
        Write-Host $doc
    }
}
else
{
    Write-Host -Object "No file with name `"$fileName`" found."
}
