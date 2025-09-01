<#
.SYNOPSIS
    Deletes a specific file if it exists and then zips the contents of the containing folder.

.DESCRIPTION
    This script first checks for the existence of a target file. If found, it deletes it.
    It then compresses the contents of the folder that held the file into a .zip archive.

.PARAMETER TargetFile
    The full path to the file you want to delete. This is a mandatory parameter.
    Example: "C:\Users\YourUser\Documents\ProjectA\temp.log"

.PARAMETER ZipDestinationPath
    The full path for the output .zip file. This is a mandatory parameter.
    Example: "C:\Users\YourUser\Desktop\ProjectA_archive.zip"

.EXAMPLE
    .\YourScript.ps1 -TargetFile "C:\data\reports\old_report.docx" -ZipDestinationPath "C:\archives\reports.zip"

    This command will delete the file 'old_report.docx' from 'C:\data\reports\' if it exists,
    and then create a zip file named 'reports.zip' in 'C:\archives\' containing everything
    that was in the 'C:\data\reports\' folder.
param (
    [Parameter(Mandatory=$true, HelpMessage="Enter the full path to the file you want to delete.")]
    [string]$TargetFile,

    [Parameter(Mandatory=$true, HelpMessage="Enter the full path for the output zip file (e.g., C:\archives\backup.zip).")]
    [string]$ZipDestinationPath
)
#>

$TargetFile = "C:\Users\pol_m\Desktop\44YORK_INST-NUI\custom1.bak"
$ZipDestinationPath = "C:\Users\pol_m\Desktop\44YORK-NUI.zip"


# --- Main Script ---

# Step 1: Check if the target file exists
if (Test-Path -Path $TargetFile -PathType Leaf) {
    Write-Host "File '$TargetFile' found. Attempting to delete..."
    try {
        # Step 2: Delete the file if it exists
        Remove-Item -Path $TargetFile -Force -ErrorAction Stop
        Write-Host "File deleted successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to delete the file '$TargetFile'. Error: $_"
        # Exit the script as we don't want to proceed if deletion fails
        return
    }
}
else {
    Write-Host "File '$TargetFile' not found. Skipping deletion." -ForegroundColor Yellow
}

# Step 3: Get the path of the folder that contained the file
$ContainingFolder = Split-Path -Path $TargetFile -Parent

# Verify the containing folder exists before trying to zip it
if (-not (Test-Path -Path $ContainingFolder -PathType Container)) {
    Write-Error "The containing folder '$ContainingFolder' does not exist. Cannot proceed."
    return
}

Write-Host "Preparing to zip the contents of folder: '$ContainingFolder'"

# Step 4: Zip the contents of the containing folder
try {
    # Using "$ContainingFolder\*" ensures that the contents of the folder are zipped,
    # rather than the folder itself being the top-level item in the zip.
    # The -Force parameter will overwrite the destination zip file if it already exists.
    Compress-Archive -Path "$ContainingFolder\*" -DestinationPath $ZipDestinationPath -Force -ErrorAction Stop
    Write-Host "Folder successfully zipped to '$ZipDestinationPath'." -ForegroundColor Green
}
catch {
    Write-Error "Failed to create the zip archive. Error: $_"
}