<#
.SYNOPSIS
    (Updated with Diagnostics) Analyzes an Excel sheet to find duplicate reading lists.
#>

# --- CONFIGURATION ---
$filePath = "C:\Users\pol_m\Downloads\duplicate reading lists august 2025 (2).xlsx" 

# --- SCRIPT ---

# 1. Check for and Import the necessary module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The 'ImportExcel' module is not installed." -ForegroundColor Yellow
    Write-Host "Please run: Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    return # Stop the script
}
Import-Module ImportExcel

# 2. Check if the file exists before proceeding
if (-not (Test-Path $filePath)) {
    Write-Host "Error: The file was not found at '$filePath'." -ForegroundColor Red
    Write-Host "Please update the '`$filePath' variable in the script." -ForegroundColor Red
    return # Stop the script
}

Write-Host "Processing file: $filePath" -ForegroundColor Green

try {
    # --- NEW DIAGNOSTIC STEP ---
    # 3. Get information about all sheets in the Excel file first
    Write-Host "Checking for available worksheets..."
    $sheetInfo = Get-ExcelSheetInfo -Path $filePath
    $sheetNames = $sheetInfo.Name
    
    if ($null -eq $sheetNames) {
        Write-Host "Error: Could not find ANY worksheets in the file. The file might be empty or corrupt." -ForegroundColor Red
        return
    }

    Write-Host "Found the following worksheets: $($sheetNames -join ', ')" -ForegroundColor Cyan

    # Check if the 'full' sheet exists in the list of names found
    if ('full' -notin $sheetNames) {
        Write-Host "Error: A worksheet named exactly 'full' was not found." -ForegroundColor Red
        Write-Host "Please check for typos, extra spaces, or case sensitivity." -ForegroundColor Red
        return
    }
    # --- END DIAGNOSTIC STEP ---


    # 4. Import the data from the 'full' sheet
    Write-Host "Importing data from the 'full' sheet..."
    $data = Import-Excel -Path $filePath -WorksheetName 'full'

    # 5. Add a specific check to see if the import returned null
    if ($null -eq $data) {
        Write-Host "Error: The script successfully found the 'full' sheet, but it appears to be empty." -ForegroundColor Red
        Write-Host "Please ensure there is data in the sheet and try again." -ForegroundColor Red
        return
    }

    # 6. Group by Course Code and Reading List Name, then find groups with more than 1 item
    Write-Host "Finding duplicate records..."
    $duplicateGroups = $data | Group-Object -Property 'Course Code', 'Reading List Name' | Where-Object { $_.Count -gt 1 }

    # 7. For each group of duplicates, find the one with the oldest modification date
    $oldestDuplicates = foreach ($group in $duplicateGroups) {
        $group.Group | Sort-Object -Property @{Expression={[datetime]$_.'Reading List Modification Date'}} | Select-Object -First 1
    }

    # 8. Check if any duplicates were found and export to the 'dupes' sheet
    if ($null -ne $oldestDuplicates -and $oldestDuplicates.Count -gt 0) {
        Write-Host "Found $($oldestDuplicates.Count) records to write to the 'dupes' sheet." -ForegroundColor Green
        
        $outputData = $oldestDuplicates | Select-Object 'Course Code', 'Course ID','Academic Department', 'Reading List ID', 'Reading List Name', 'Reading List Modification Date'
        
        Write-Host "Exporting results to the 'dupes' sheet..."
        $outputData | Export-Excel -Path $filePath -WorksheetName 'dupes' -AutoSize -ClearSheet
        
        Write-Host "Script finished successfully. The 'dupes' sheet has been created/updated in your file." -ForegroundColor Green
    }
    else {
        Write-Host "No duplicate records found matching the criteria." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An unexpected error occurred during script execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}