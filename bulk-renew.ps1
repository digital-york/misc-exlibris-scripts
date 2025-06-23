# PS script to read through file of user ids/loans
# and renew
#**REMOVE API VALUE BEFORE SUBMITTING**

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-eu.hosted.exlibrisgroup.com/"
$queryParams = '?' + 'op=renew' + '&' + [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('API KEY GOES HERE');

$file = "C:\Work\renewals\renew.xlsx"
$sheetName = "renew"

#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

#Count max row
$rowMax = ($sheet.UsedRange.Rows).count

Write-Host "We have " ($rowMax-1) " rows to process"

#Declare the starting positions
$rowusr,$coluser,$colid = 1,1,2


#loop through sheet and act on each row
	for ($i=0; $i -le $rowMax-1; $i++)
	{

	   $user = $sheet.Cells.Item($rowusr+$i,$coluser).text

       # $user = $sheet.Cells.Item(1,1).text

        $loan_id = $sheet.Cells.Item($rowusr+$i,$colid).text
		
		#test for empty user id
		if (!$user){
			Write-Host "Null user id. Exiting."
			Exit
		}
		
		Write-Host "Processing user" $user "loan id " $loan_id
			
        #renew loan
        $renew_url = $url_prefix + "almaws/v1/users/" + $user + "/loans/" + $loan_id +  $queryParams

        Write-Host ($renew_url)
		

        #renew loan	
	try{		
			Invoke-RestMethod -Method 'POST' -Uri $renew_url -ContentType 'application/xml' -Body $xml
			#Start-Sleep 3
		}
	catch	
		{
		write-host "Fatal error: Renew" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        $error_dets = $user + ',' + $loan_id
		Add-Content "C:\Work\renewals\renew-errors.txt" $error_dets
		}				
	}
#close excel file
$objExcel.quit()

Write-Host "Processing Complete"
