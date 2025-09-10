<#
this script is intended to identify courses in the source spreadsheet with duplicate reading lists.
If a dupe is found, the relevant details of the oldest modification date are saved to a new tab for future processing
#>

# --- CONFIGURATION ---
$filePath = "C:\Users\pol_m\Downloads\duplicate reading lists august 2025 (3).xlsx" 

# --- SCRIPT ---

# 1. Check for and Import the necessary module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The 'ImportExcel' module is not installed." -ForegroundColor Yellow
    Write-Host "Please run: Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    return
}
Import-Module ImportExcel

# 2. Check if the file exists before proceeding
if (-not (Test-Path $filePath)) {
    Write-Host "Error: The file was not found at '$filePath'." -ForegroundColor Red
    return
}

Write-Host "Processing file: $filePath" -ForegroundColor Green

try {
    # 3. Import the data from the 'full' sheet
    Write-Host "Importing data from the 'full' sheet..."
    $data = Import-Excel -Path $filePath -WorksheetName 'full'

    if ($null -eq $data) {
        Write-Host "Error: The sheet 'full' appears to be empty." -ForegroundColor Red
        return
    }

    # 4. Group by Course Code and Reading List Name
    Write-Host "Finding duplicate records..."
    $duplicateGroups = $data | Group-Object -Property 'Course Code', 'Reading List Name' | Where-Object { $_.Count -gt 1 }

    # 5. For each group of duplicates, find the one with the oldest modification date
    $oldestDuplicates = foreach ($group in $duplicateGroups) {
        
        $oldestRecordSoFar = $null
        $oldestDateSoFar = $null

        foreach ($record in $group.Group) {
            # --- CRITICAL FIX: Force the value to a string before processing ---
            $dateObject = $record.'Reading List Modification Date'
            if ($null -eq $dateObject) { continue } # Skip if the property itself is null

            $dateString = $dateObject.ToString().Trim()
            if ([string]::IsNullOrWhiteSpace($dateString)) { continue }

            # Diagnostic line to see what the script is parsing
            Write-Host "DEBUG: Attempting to parse date string: '$dateString'"

            $currentDate = $null
            # Manually parse the different string formats to bypass locale issues
            if ($dateString -match '(\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2})') {
                $currentDate = New-Object datetime($matches[3], $matches[2], $matches[1], $matches[4], $matches[5], 0)
            } elseif ($dateString -match '(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})') {
                $currentDate = New-Object datetime($matches[1], $matches[2], $matches[3], $matches[4], $matches[5], $matches[6])
            }

            # If the date was successfully parsed, compare it
            if ($currentDate -ne $null) {
                if ($oldestRecordSoFar -eq $null -or $currentDate -lt $oldestDateSoFar) {
                    $oldestDateSoFar = $currentDate
                    $oldestRecordSoFar = $record
                }
            }
        }
        # After checking all records in the group, output the oldest one found
        if ($null -ne $oldestRecordSoFar) {
            $oldestRecordSoFar
        }
    }

    # 6. Check if any duplicates were found and export to the 'dupes' sheet
    $validOldestDuplicates = $oldestDuplicates | Where-Object { $_ -ne $null }
    if ($validOldestDuplicates.Count -gt 0) {
        Write-Host "Found $($validOldestDuplicates.Count) records to write to the 'dupes' sheet." -ForegroundColor Green
        
        $outputData = $validOldestDuplicates | Select-Object 'Course Code', 'Course ID','Academic Department', 'Reading List ID', 'Reading List Name', 'Reading List Modification Date'
        
        Write-Host "Exporting results to the 'dupes' sheet..."
        $outputData | Export-Excel -Path $filePath -WorksheetName 'dupes' -AutoSize -ClearSheet
        
        Write-Host "Script finished successfully. The 'dupes' sheet has been created/updated in your file." -ForegroundColor Green
    }
    else {
        Write-Host "No duplicate records found with valid, parsable dates." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An unexpected error occurred during script execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}