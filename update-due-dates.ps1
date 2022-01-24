# PS script to read through file of barcodes ...
# ... retrieve item link for each and perform scan-in
# uses BIBS API (Bulk Returns)

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-na.hosted.exlibrisgroup.com/"
$queryParams = '?' +  [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('l8xx1f7d2163f0884b06b952e3f942e80b47');

#$api_key = [System.Web.HttpUtility]::UrlEncode($key) 
$file = "C:\Work\overdue Lost Loans\test.xlsx"
$sheetName = "overdue"
$library = "EXTST"
$circ_desk = "DEFAULT_CIRC_DESK"

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
	for ($i=1; $i -le $rowMax-1; $i++)
	{

		$user = $sheet.Cells.Item($rowusr+$i,$coluser).text

        $loan_id = $sheet.Cells.Item($rowusr+$i,$colid).text
		
		#test for empty user id
		if (!$user){
			Write-Host "Null user id. Exiting."
			Exit
		}
		
		Write-Host "Processing user" $user " row number" ($rowusr) "loan id " $loan_id
		#retrieve item info
		$loan_dets = $url_prefix + "almaws/v1/users/" + $user + "/loans/" + $loan_id + $queryParams
		
        #https://api-eu.hosted.exlibrisgroup.com/almaws/v1/users/cc1684/loans/42358939650001381?apikey=l8xx1f7d2163f0884b06b952e3f942e80b47


		Write-Host "Loan Details URL " $loan_dets
			
	try{		
			#API call and assign result to xml variable
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $loan_dets	
			#Start-Sleep 3
		}
	
	catch	
		{
		write-host "Fatal error: Get Loan Details" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
		Add-Content "C:\Work\Bulk returns\get-item-errors.txt" $barcode
		}
				
		

		$due_date = $xml.item_loan.due_date
        $mms_id = $xml.item_loan.mms_id
        $holding_id = $xml.item_loan.holding_id
        $item_id = $xml.item_loan.item_id
        $loan_id = $xml.item_loan.loan_id
			
		Write-Host "Due date is"  $due_date

        $xml.item_loan.due_date = "2022-10-01T22:59:00Z"
        
        $update_url = $url_prefix + "almaws/v1/users/" + $user + "/loans/" + $loan_id +  $queryParams

        Write-Host $update_url


        #update due date
        
        Invoke-RestMethod -Method 'PUT' -Uri $update_url -ContentType 'application/xml' -Body $xml
      

				
	}
#close excel file
$objExcel.quit()

Write-Host "Processing Complete"
