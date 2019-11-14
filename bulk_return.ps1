# PEH Sept 2018
# PS script to read through file of barcodes ...
# ... retrieve item link for each and perform scan-in
# uses BIBS API (Bulk Returns)

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-na.hosted.exlibrisgroup.com/"
$queryParams = '&' +  [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('l7xxc7a91dcc39b64ba0b9a40fd374e559e1');
#$key = "l7xxc7a91dcc39b64ba0b9a40fd374e559e1"
#$api_key = [System.Web.HttpUtility]::UrlEncode($key) 
$file = "C:\Work\Alma\returns.xlsx"
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
	
		$item_by_barc_url = $url_prefix + "almaws/v1/items?item_barcode=" + $barcode + $queryParams
			
	try{		
			#API call and assign result to xml variable
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $item_by_barc_url	
			Start-Sleep 5
		}
	
	catch	
		{
		write-host "Fatal error: Get Item By Barcode" -ForegroundColor Red
		write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
		write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
		Add-Content "C:\Work\Alma\Scripts\Errors\get-item-errors.txt" $barcode
		}
				
		#scan through returned xml and retrieve item link (redirected to from the barcode url)
		#this is required to carry out the scan-in operation
		$item_url = $xml.item.link
		
	
		#construct scan-in link
		$scan_in_url = $item_url +  "?op=scan&library=" + $library + "&circ_desk=" + $circ_desk + $queryParams
		
	try{
			#scan item in
			Invoke-WebRequest -Method 'POST' -Uri $scan_in_url -ContentType application/xml 
			Start-Sleep 5
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



