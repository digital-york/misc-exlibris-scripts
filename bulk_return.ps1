# Updated PEH 2021 - perform check for active holds 
#PEH Sept 2018
# PS script to read through file of barcodes ...
# ... retrieve item link for each and perform scan-in
# uses BIBS API (Bulk Returns)

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-na.hosted.exlibrisgroup.com/"
$queryParams = '&' +  [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('**API key here**');

#$api_key = [System.Web.HttpUtility]::UrlEncode($key) 
$file = "C:\Work\Bulk returns\returns.xlsx"
$sheetName = "results"
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
$rowbarc,$colbarc = 1,1


#loop through sheet and act on each row
	for ($i=1; $i -le $rowMax-1; $i++)
	{
		$barcode = $sheet.Cells.Item($rowbarc+$i,$colbarc).text
		
		#test for empty barcode
		if (!$barcode){
			Write-Host "Null barcode. Exiting."
			Exit
		}
		
		Write-Host "Processing barcode" $barcode " row number" ($rowbarc+$i) 
		#retrieve item info
		$item_by_barc_url = $url_prefix + "almaws/v1/items?item_barcode=" + $barcode + $queryParams
		
		#Write-Host "Item URL " $item_by_barc_url
			
	try{		
			#API call and assign result to xml variable
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $item_by_barc_url	
			Start-Sleep 3
		}
	
	catch	
		{
		write-host "Fatal error: Get Item By Barcode" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
		Add-Content "C:\Work\Bulk returns\get-item-errors.txt" $barcode
		}
				
		#scan through returned xml and retrieve item link (redirected to from the barcode url)
		#this is required to carry out the scan-in operation
		$item_url = $xml.item.link
			
		Write-Host $item_url
		
		#gather parameters for outstanding requests
		$holding_id = $xml.item.holding_data.holding_id
		$itemid = $xml.item.item_data.pid
		$mmsid = $xml.item.bib_data.mms_id
		
		
		#construct request check link here, return request data
		$req_url = $item_url + "/requests?" + $queryParams
		
		[xml]$req_xml = Invoke-RestMethod -Method 'GET' -Uri $req_url 	
		
		Start-Sleep 3
		
		$req_count = $req_xml.user_requests.total_record_count	
			
		if ($req_count -ne '0'){
		
			Write-Host "Request Found"
			Add-Content "C:\Work\Bulk returns\items-with-requests.txt" $barcode
		
		}
		
		Write-Host $req_url

		#construct scan-in link
		$scan_in_url = $item_url +  "?op=scan&library=" + $library + "&circ_desk=" + $circ_desk + $queryParams
		
	try{
			#scan item in, but only if there isn't an active hold
			
			if ($req_count -ne '0'){
				}else{
					Write-Host "Returning Item"
					Invoke-WebRequest -Method 'POST' -Uri $scan_in_url -ContentType application/xml 
			    }
			Start-Sleep 3
		}
	catch
		{
			write-host "Fatal error: Scan In" -ForegroundColor Red
			write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
			write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
			Add-Content "C:\Work\Alma\Scripts\Errors\scan-in-errors.txt" $barcode			
		}
				
	}
#close excel file
$objExcel.quit()

Write-Host "Processing Complete"
