<#
.SYNOPSIS
    (Module-Free Version) Reads a CSV file to find duplicate reading lists for courses,
    identifies the OLDEST record in each duplicate group, and constructs an API URL.
    This script uses built-in PowerShell commands and requires no external modules.

.DESCRIPTION
    1.  Sets the path to the source CSV file.
    2.  Imports the data using the built-in 'Import-Csv'. This correctly handles long IDs as text.
    3.  Groups the data by 'Course Code' and 'Reading List Name'.
    4.  Filters to find only the groups with more than one item (the duplicates).
    5.  For each group, it explicitly converts the modification date string into a proper DateTime object
        to ensure accurate, chronological sorting.
    6.  It sorts the group by this date and selects the oldest record.
    7.  It then constructs the final API URL from the 'Course ID' and 'Reading List Id' of that oldest record.
    8.  All generated URLs are printed to the console.
#>

# --- CONFIGURATION ---
# !!! IMPORTANT: Update this to the full path or just the filename if it's in the same folder as the script.
$csvFilePath = "C:\Work\dupes.csv"

# --- SCRIPT ---

# 1. Check if the source file exists
if (-not (Test-Path $csvFilePath)) {
    Write-Host "Error: The source CSV file was not found at '$csvFilePath'." -ForegroundColor Red
    Write-Host "Please update the '`$csvFilePath' variable in the script." -ForegroundColor Red
    return # Stop the script
}

Write-Host "Processing file: $csvFilePath" -ForegroundColor Green

try {
    # 2. Import data using the built-in Import-Csv command. No extra modules needed.
    $data = Import-Csv -Path $csvFilePath

    # 3. Group by Course Code and Reading List Name to find duplicates
    # The '.GetEnumerator()' is used to handle potential single-item groups gracefully.
    $duplicateGroups = $data | Group-Object -Property 'Course Code', 'Reading List Name' | Where-Object { $_.Count -gt 1 }

    if ($null -eq $duplicateGroups) {
        Write-Host "No duplicate records found matching the criteria." -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($duplicateGroups.Count) groups of duplicates. Identifying the oldest record in each..." -ForegroundColor Cyan
    Write-Host "--- Generating URLs ---"

    # 4. Iterate through each group of duplicates
foreach ($group in $duplicateGroups) {
        # 5. Find the oldest record in the group
        $oldestRecord = $group.Group |
            # --- FIX APPLIED HERE ---
            # First, filter out any rows where the modification date is null or just whitespace.
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.'Reading List Modification Date') } |
            # Now, it is safe to sort the remaining (valid) rows.
            Sort-Object -Property @{Expression={ [datetime]$_.'Reading List Modification Date' }} |
            Select-Object -First 1

        # Check if we found a valid oldest record (it could be null if all dates in the group were blank)
        if ($null -ne $oldestRecord) {
            # 6. Construct the URL from the oldest record's data
            $courseId = $oldestRecord.'Course ID'
            $readingListId = $oldestRecord.'Reading List Id'

            if (-not [string]::IsNullOrWhiteSpace($courseId) -and -not [string]::IsNullOrWhiteSpace($readingListId)) {
                $url = "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/courses/$courseId/reading-lists/$readingListId"
                Write-Host $url
            }
            else {
                Write-Host "Warning: Skipping a record in group '$($group.Name)' due to a missing ID." -ForegroundColor Yellow
            }
        }
        else {
             Write-Host "Warning: Skipping group '$($group.Name)' because no valid modification dates were found." -ForegroundColor Yellow
        }
    }

    Write-Host "--- Script finished ---" -ForegroundColor Green
}
catch {
    Write-Host "An unexpected error occurred during script execution:" -ForegroundColor Red
    # This will print the specific error message, e.g., about date formatting
    Write-Host $_.Exception.Message -ForegroundColor Red
}