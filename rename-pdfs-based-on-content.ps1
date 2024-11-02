# Set the location where the PDF files are stored
Set-Location "C:\Path\To\Your\Files"

# Add the iTextSharp library for PDF manipulation
Add-Type -Path "C:\Path\To\iTextSharp\itextsharp.dll"

# Get all PDF files in the directory and subdirectories
$allFiles = Get-ChildItem -Recurse -Include "*.pdf" -File | Select Name, Length, FullName

# Export file details to a CSV file
$csvFile = "$PSScriptRoot\allFiles.csv"
$allFiles | Export-Csv -Path $csvFile -NoTypeInformation

# Define invalid characters for file names
$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''

# Initialize an array to store renamed file details
$exportArray = @()

$rowCount = 0
foreach ($row in $allFiles) {
    $exportItem = "" | Select "oldName", "newName"
    $PDF = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $row.FullName

    # Determine new file name based on content
    if ($row.Name -like "*Feedback Form*") {
        $secondPageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PDF, 2)
        $textParsing = $secondPageText.Split("Project or Initiative Name ")[1]
        $fileNameAppendText = $textParsing.Split("Select the type of work")[0].Trim()
        $newFileNameBase = ($row.Name -split "_")[0..3] -join "_"
    }
    elseif ($row.Name -like "*Project Reviews*") {
        $firstPageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PDF, 1)
        $fileNameParsing = $firstPageText.Split("Review Period")[0].Trim()
        $fileNameParsing = $fileNameParsing.Replace(">", "GT").Replace("<", "LT").Replace("`n", " ").Replace("/", " or ")
        $userIDAndName = ($row.Name -split "_")[0..2] -join "_"
        $newFileNameBase = $userIDAndName + "_" + $fileNameParsing

        if ($fileNameParsing -like "* GT *") {
            $secondPageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PDF, 2)
            $projectNameParsing = $secondPageText -Split "`n"
            $fileNameAppendText = $projectNameParsing[4]
        }
        elseif ($fileNameParsing -like "* LT *") {
            $searchString = if ($fileNameParsing -like "*(Project Manager or Team Lead Initiated)*") {
                "( Review Creator )"
            } elseif ($fileNameParsing -like "*(Employee Initiated)*") {
                "( Project Manager ) "
            }

            $secondPageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PDF, 2)
            $replaceRegexSearchString = "\" + $searchString.Replace(")", "\)") + ".*"
            $likeSearchString = "*" + $searchString + "*"
            $projectManagerCheckArray = ($secondPageText -Split "`n" | Where-Object { $_ -like $likeSearchString })

            if ($null -eq $projectManagerCheckArray) {
                Write-Host "No Project Manager listed for $($row.Name)"
                $projectManagerName = ""
            }
            else {
                $projectManagerName = $projectManagerCheckArray[0] -Replace $replaceRegexSearchString
            }
            $fileNameAppendText = $projectManagerName
        }
    }

    # Construct the new file name
    $newName = $newFileNameBase + "_" + $fileNameAppendText + ".pdf"
    $newName = $newName.Replace("`n", " ").Replace("/", "-").Replace("\", "-").-replace "[${invalidChars}]", '_'

    # Ensure the new name is unique
    $dupeCount = 1
    While ($exportArray.newName -contains $newName) {
        $newName = $newFileNameBase + "_" + $fileNameAppendText + "(" + $dupeCount + ").pdf"
        $newName = $newName.Replace("`n", " ").Replace("/", "-").Replace("\", "-").-replace "[${invalidChars}]", '_'
        $dupeCount++
    }

    # Dispose of the PDF object to release the file lock
    $PDF.Dispose()

    # Update the last modified date and rename the file
    (Get-Item $row.FullName).LastWriteTime = (Get-Date)
    try {
        Rename-Item -Path $row.FullName -NewName $newName
    }
    catch {
        Write-Host $row.FullName "|" $newName -ForegroundColor Red
    }

    # Increment row count and display progress every 100 files
    $rowCount++
    if ($rowCount % 100 -eq 0) {
        Write-Host "$rowCount rows processed out of $($allFiles.Count)"
    }

    # Set export item for logging
    $exportItem.oldName = $row.Name
    $exportItem.newName = $newName
    $exportArray += $exportItem

    # Display progress
    $percentComplete = ($rowCount / $allFiles.Count) * 100
    Write-Progress -Activity "Renaming Files" -Status "Processing..." -PercentComplete $percentComplete
}

# Export log of changes
$ISODate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")
$exportArray | Export-Csv -Path ".\$ISODate File Name Changes.csv" -Force -Encoding Unicode

# Use 7Zip to compress the entire folder of renamed PDFs
$7ZipExecutable = "C:\Program Files\7-Zip\7z.exe"
& $7ZipExecutable a -tzip HistFeedbackProjectReviewsRenamed.zip *.pdf
