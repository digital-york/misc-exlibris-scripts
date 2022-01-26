# PS script to read through file of user ids/loans
# ... apply 14 day extension 

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-eu.hosted.exlibrisgroup.com/"
$queryParams = '?' + [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('**API KEY GOES HERE');
$file = "C:\Work\overdue Lost Loans\overdue.xlsx"
$sheetName = "overdue"

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
		
		Write-Host "Processing user" $user "loan id " $loan_id
		#retrieve item info
		$loan_dets = $url_prefix + "almaws/v1/users/" + $user + "/loans/" + $loan_id + $queryParams
		

		Write-Host "Loan Details URL " $loan_dets
			
	try{		
			#API call and assign result to xml variable
            #retrieve loan detail xml
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $loan_dets	
			#Start-Sleep 3
		}
	
	catch	
		{
		write-host "Fatal error: Get Loan Details" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
		Add-Content "C:\Work\overdue Lost Loans\get-item-errors.txt" $loan_id
		}


		$due_date = $xml.item_loan.due_date
        $mms_id = $xml.item_loan.mms_id
        $holding_id = $xml.item_loan.holding_id
        $item_id = $xml.item_loan.item_id
        $loan_id = $xml.item_loan.loan_id
			       
        #set new due date     
        $xml.item_loan.due_date = '2022-02-07T23:59:00'
        
        $update_url = $url_prefix + "almaws/v1/users/" + $user + "/loans/" + $loan_id +  $queryParams

        #update due date		
	try{		
			Invoke-RestMethod -Method 'PUT' -Uri $update_url -ContentType 'application/xml' -Body $xml
			Start-Sleep 3
		}
	catch	
		{
		write-host "Fatal error: Update Due Date" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
		Add-Content "C:\Work\overdue Lost Loans\update-errors.txt" $loan_id
		}				
	}
#close excel file
$objExcel.quit()

Write-Host "Processing Complete"
