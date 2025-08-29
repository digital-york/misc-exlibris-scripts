<#
.SYNOPSIS
    Downloads all .mrc files from the University of Hull Library API.

.DESCRIPTION
    This script connects to the specified MARC records API URL, parses the HTML to find all links
    ending in .mrc, and then downloads each file to a user-specified destination folder.
    It will skip any files that have already been downloaded.

.AUTHOR
    Gemini
#>

# --- Configuration ---
$sourceUrl = "https://api.library.hull.ac.uk/marc_records/"

# --- Main Script ---

# 1. Get the destination folder from the user
try {
    $destinationFolder = Read-Host -Prompt "Please enter the full path for the destination folder (e.g., C:\MARC_Downloads)"
    
    if ([string]::IsNullOrWhiteSpace($destinationFolder)) {
        Write-Warning "No destination folder provided. Exiting script."
        return # Exit the script
    }

    # 2. Check if the destination folder exists. If not, create it.
    if (-not (Test-Path -Path $destinationFolder -PathType Container)) {
        Write-Host "Destination folder '$destinationFolder' does not exist. Creating it now..."
        # The -Force switch creates any necessary parent directories as well.
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
    } else {
        Write-Host "Using existing destination folder: '$destinationFolder'"
    }

    # 3. Get the list of files from the web page
    Write-Host "Fetching list of files from $sourceUrl ..."
    $response = Invoke-WebRequest -Uri $sourceUrl -UseBasicParsing
    
    # Filter the links to get only the ones that end with '.mrc'
    $fileLinks = $response.Links | Where-Object { $_.href -like '*.tar.gz' }

    if ($null -eq $fileLinks -or $fileLinks.Count -eq 0) {
        Write-Warning "Could not find any '.mrc' files at the source URL. Exiting."
        return
    }

    $totalFiles = $fileLinks.Count
    Write-Host "Found $totalFiles files to process."
    Write-Host "--------------------------------------------------"

    # 4. Loop through each link and download the file
    $fileCounter = 0
    foreach ($link in $fileLinks) {
        $fileCounter++
        $fileName = $link.href
        $fullSourcePath = "$sourceUrl$fileName"
        $fullDestinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName

        # Check if the file already exists before downloading
        if (Test-Path -Path $fullDestinationPath) {
            Write-Host "($fileCounter/$totalFiles) SKIPPING: '$fileName' already exists."
        } else {
            Write-Host "($fileCounter/$totalFiles) DOWNLOADING: '$fileName'..."
            try {
                # Use Invoke-WebRequest to download the file
                Invoke-WebRequest -Uri $fullSourcePath -OutFile $fullDestinationPath
                Write-Host " -> Successfully saved to '$fullDestinationPath'" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to download '$fileName'. Error: $($_.Exception.Message)"
            }
        }
    }

    Write-Host "--------------------------------------------------"
    Write-Host "Download process complete." -ForegroundColor Cyan
}
catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
}