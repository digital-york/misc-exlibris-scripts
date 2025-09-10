<#
This script scans through a list of courses for duplicate reading lists.
If a duplicate exists, it identifies the version with the oldest modification date
and removes this via the Alma api

#>

# --- CONFIGURATION ---
# source file
$filePath = "C:\Work\dupes.csv"
$apiKey = "***INSERT API KEY VALUE***"
$outputLogFile = "C:\Work\log.txt"
# --- SCRIPT ---

# 1. Check if the source file exists
if (-not (Test-Path $filePath)) {
    Write-Host "Error: The source file was not found at '$filePath'." -ForegroundColor Red
    return
}

Write-Host "Processing file: $filePath (as a Tab-Separated file)..." -ForegroundColor Green

try {
    # 2. Import the data
    $allData = Import-Csv -Path $filePath -Delimiter "`t"

    # 3. Filter for valid records.
    $validData = $allData | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.'Course Code') -and
        -not [string]::IsNullOrWhiteSpace($_.'Reading List Name')
    }

    if ($validData.Count -eq 0) {
        Write-Host "No valid data rows found after filtering." -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($validData.Count) valid data rows to process..."

    # 4. Group the clean data to find duplicates
    $duplicateGroups = $validData | Group-Object -Property 'Course Code', 'Reading List Name' | Where-Object { $_.Count -gt 1 }

    if ($duplicateGroups.Count -eq 0) {
        Write-Host "No duplicate records were found in the valid data." -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($duplicateGroups.Count) groups of duplicates. Generating URLs..."
    Write-Host "--- Generating URLs ---"

    # 5. Iterate through each group and find the oldest record
    foreach ($group in $duplicateGroups) {
        $sortedRecords = $group.Group |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.'Reading List Modification Date') } |
            Select-Object *, @{Name="ParsedDate"; Expression={
                $dateString = $_.'Reading List Modification Date'.Trim()
                # Match different date patterns and construct the datetime object manually.
                if ($dateString -match '(\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2})') { # dd/MM/yyyy HH:mm
                    # $matches[3] = Year, $matches[2] = Month, $matches[1] = Day, etc.
                    return New-Object datetime($matches[3], $matches[2], $matches[1], $matches[4], $matches[5], 0)
                } elseif ($dateString -match '(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})') { # yyyy-MM-dd HH:mm:ss
                    return New-Object datetime($matches[1], $matches[2], $matches[3], $matches[4], $matches[5], $matches[6])
                } else {
                    Write-Host "Warning: Could not recognize the date format for '$dateString'. Skipping." -ForegroundColor Yellow
                    return $null
                }
            }} |
            Where-Object { $_.ParsedDate -ne $null } |
            Sort-Object -Property ParsedDate

        if ($sortedRecords) {
            $oldestRecord = $sortedRecords | Select-Object -Last 1

            # 6. Construct the URL
            $courseId = $oldestRecord.'Course ID'
            $readingListId = $oldestRecord.'Reading List Id'

            if (-not [string]::IsNullOrWhiteSpace($courseId) -and -not [string]::IsNullOrWhiteSpace($readingListId)) {
                $url = "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/courses/$($courseId.Trim())/lists/$($readingListId.Trim())?apikey=$($apiKey)"
                Write-Host $url
                $courseCode = $oldestRecord.'Course Code'
                $listName = $oldestRecord.'Reading List Name'
                $modDate = $oldestRecord.'Reading List Modification Date'

# log results
$courseCode = $oldestRecord.'Course Code'
$listName = $oldestRecord.'Reading List Name'
$modDate = $oldestRecord.'Reading List Modification Date'

# Create a formatted string for the log file
$logOutput = @"
URL: $url
Course Code: $courseCode
Reading List Name: $listName
Modification Date: $modDate
----------------------------------
$logOutput = @"
URL: {0}
Course Code: {1}
Reading List Name: {2}
Modification Date: {3}
----------------------------------
"@ -f $url, $courseCode, $listName, $modDate

# Append the string to the log file
Add-Content -Path $outputLogFile -Value $logOutput
Write-Host "--> Wrote URL for list '$listName' to $outputLogFile" -ForegroundColor Cyan
Write-Host "---"
                
# --- API ACTION ---
        Write-Host "Attempting to DELETE Reading List ID $($readingListId.Trim())..." -ForegroundColor Yellow
        try {
            # To run for real, remove the '#' from the beginning of the next line.
            #Invoke-RestMethod -Uri $url -Method Delete

            Write-Host "SUCCESS: The DELETE request was sent." -ForegroundColor Green
        }
        catch {
            $errorMessage = $_.Exception.Response.StatusCode
            Write-Host "FAILED: The API returned an error: $errorMessage" -ForegroundColor Red
        }
        Write-Host "---"
    }
 }
}

    Write-Host "Script finished. Log file created at: $outputLogFile" -ForegroundColor Green
}
catch {
    Write-Host "An unexpected error occurred during script execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}