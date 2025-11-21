<#
.SYNOPSIS
    Compares two CSV files and finds rows in File A that do not exist in File B,
    based on user-specified key columns.
#>

# --- 1. Get User Input ---

Write-Host "--- Spreadsheet Comparison Tool ---" -ForegroundColor Yellow

#$fileA = Read-Host "Enter the full path to the first spreadsheet (File A)"
#$fileB = Read-Host "Enter the full path to the second spreadsheet (File B)"
#$keyColumnsPrompt = Read-Host "Enter the key column(s) to compare (e.g., 'ID' or 'Email,LastName')"

$fileA = "C:\Work\Leganto\alma.xlsx"
$fileB = "C:\Work\Leganto\hyms.xlsx"
# --- 2. Validate Inputs ---

if (-not (Test-Path $fileA)) {
    Write-Error "Error: File A not found at path: $fileA"
    return # Stop script execution
}
if (-not (Test-Path $fileB)) {
    Write-Error "Error: File B not found at path: $fileB"
    return # Stop script execution
}

# Split the column names string into an array and trim whitespace
#$keyColumns = $keyColumnsPrompt -split ',' | ForEach-Object { $_.Trim() }

$keyColumns = "Name"

if ($keyColumns.Count -eq 0 -or ($keyColumns.Count -eq 1 -and [string]::IsNullOrWhiteSpace($keyColumns[0]))) {
     Write-Error "Error: No valid key columns were provided."
     return
}

Write-Host "Comparing files based on column(s): $($keyColumns -join ', ')" -ForegroundColor Cyan

# --- 3. Import Data ---

try {
    Write-Host "Importing $fileA..."
    # Using -ErrorAction Stop to ensure the 'catch' block triggers on import failure
    $dataA = Import-Csv -Path $fileA -ErrorAction Stop
    
    Write-Host "Importing $fileB..."
    $dataB = Import-Csv -Path $fileB -ErrorAction Stop
} catch {
    Write-Error "Failed to import CSV files. Please check paths and file permissions."
    Write-Error "Specific error: $_"
    return
}

# --- 4. Perform Comparison ---

Write-Host "Comparing objects..."

# Compare-Object finds the differences.
# -ReferenceObject is our "source" list (File A).
# -DifferenceObject is the list to compare against (File B).
# -Property specifies which columns to check for a match.
# -PassThru outputs the *original object* that is different,
#   adding a 'SideIndicator' property.
#
# SideIndicator values:
#   '<=' : Item is only in the ReferenceObject (File A)
#   '=>' : Item is only in the DifferenceObject (File B)
#   '==' : Item is in both (only shown if -IncludeEqual is used)

$rowsOnlyInA = Compare-Object -ReferenceObject $dataA -DifferenceObject $dataB -Property $keyColumns -PassThru |
               Where-Object { $_.SideIndicator -eq '<=' } |
               Select-Object -ExcludeProperty SideIndicator

# --- 5. Display Results ---

if ($rowsOnlyInA) {
    # Use Measure-Object to get an accurate count
    $count = ($rowsOnlyInA | Measure-Object).Count
    Write-Host "Found $count row(s) in '$fileA' that do not exist in '$fileB'." -ForegroundColor Green
    
    # Display results neatly in the console
    $rowsOnlyInA | Format-Table -AutoSize
    
    # --- 6. Optional Export ---
    $outputFile = Read-Host "OPTIONAL: Enter a path to save these results (e.g., C:\temp\unique_rows.csv). Press Enter to skip"
    
    if (-not [string]::IsNullOrWhiteSpace($outputFile)) {
        try {
            # Save the results to a new CSV
            $rowsOnlyInA | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
            Write-Host "Results successfully saved to $outputFile" -ForegroundColor Cyan
        } catch {
            Write-Error "Failed to save output file: $_"
        }
    }
    
} else {
    Write-Host "All rows in '$fileA' (based on the specified columns) were found in '$fileB'." -ForegroundColor Yellow
}