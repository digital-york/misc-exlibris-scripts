# --- Configuration ---
# An array of email addresses to notify
$reportReceivers = @("paul.harding@york.ac.uk")

# The API Key (replace with your actual key)
$apiKey = "l8xx686ac87bbd61480e9dd4c2c2c4b74b53" 

# The SMTP server for sending email (replace with your server)
$smtpServer = "smtp.gmail.com" 
$emailFrom = "api-monitor@york.ac.uk"

# --- Main Script ---
# Define the headers for the request in a hashtable
$headers = @{
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
    "Authorization" = "apikey $apiKey"
}

$uri = "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/conf/test?apikey=" + $apiKey

try {
    # Make the web request.
    # -ErrorAction Stop ensures that any non-successful HTTP status code triggers the 'catch' block.
    # -ResponseHeadersVariable stores the response headers in the specified variable.
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ResponseHeadersVariable responseHeaders -ErrorAction Stop

    # Check if the required headers exist in the response
    if ($responseHeaders.ContainsKey("X-Exl-Api-Remaining") -and $responseHeaders.ContainsKey("X-Exl-Api-Quota")) {
        
        # Cast the header values to integers for calculation
        $remainingCalls = [int]$responseHeaders["X-Exl-Api-Remaining"]
        $totalQuota = [int]$responseHeaders["X-Exl-Api-Quota"]

        # Calculate the remaining percentage
        $remainingPercentage = ($remainingCalls / $totalQuota) * 100

        # Determine the subject line based on the threshold
        $subject = switch ($remainingPercentage) {
            { $_ -le 2 } { "CRITICAL: API CALLS THRESHOLD: {0:N2}% remaining" -f $remainingPercentage }
            { $_ -le 5 } { "WARNING: API CALLS THRESHOLD: {0:N2}% remaining" -f $remainingPercentage }
            default      { "INFO: API CALLS THRESHOLD: {0:N2}% remaining" -f $remainingPercentage }
        }
    }
    else {
        # This will run if the request was successful but the quota headers were missing
        $subject = "INFO: API call successful but quota headers were not found."
    }
}
catch {
    # This block runs if Invoke-RestMethod fails (e.g., 4xx or 5xx response, network error)
    $subject = "NO RESPONSE FROM API"
    Write-Warning "API call failed. Error: $($_.Exception.Message)"
}

# --- Send Notification ---
# Loop through each recipient and send the email
foreach ($receiver in $reportReceivers) {
    try {
        Send-MailMessage -To $receiver -Subject $subject -Body $subject -From $emailFrom -SmtpServer $smtpServer -ErrorAction Stop
        Write-Host "Successfully sent notification to $receiver"
    }
    catch {
        Write-Error "Failed to send email to $receiver. Error: $($_.Exception.Message)"
    }
}