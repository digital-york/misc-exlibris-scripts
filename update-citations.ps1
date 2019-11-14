# PEH Sept 2018
# PS script to remove pages from citations
# uses course API (Update Citations)

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-na.hosted.exlibrisgroup.com"
$queryParams = [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('l7xx949a811569604471b5b08164c55ba6d2');
$file = "C:\Work\Alma\citations.xlsx"
$sheetName = "Sheet1"

#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

#Count max row
$rowMax = ($sheet.UsedRange.Rows).count

Write-Host "We have " ($rowMax-1) " rows to process"

#Declare the starting positions
$rowCourse,$colCourse = 1,1
$rowList,$colList = 1,2
$rowCitation,$colCitation = 1,3

#loop to get values and store it
	for ($i=1; $i -le $rowMax-1; $i++)
	{
		$course = $sheet.Cells.Item($rowCourse+$i,$colCourse).text
		$list= $sheet.Cells.Item($rowList+$i,$colList).text
		$citation = $sheet.Cells.Item($rowCitation+$i,$colCitation).text
		
		#test for empty course code
		if (!$course){
			Write-Host "Null course. Exiting."
			Exit
		}
		
		
		#course id is required to update citation
		#this is retrieved via a courses api search using the module code
		$course_id_call = $url_prefix + "/almaws/v1/courses?q=code~" + $course + '&' + $queryParams
		
		#Write-Host "ID call is "$course_id_call
	
		#get returned xml
		[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $course_id_call 
		
		#retrieve course id from xml
		$course_id = $xml.courses.course.id

	
		#$nodes | % {
			#$course_id = $_.SelectSingleNode('//id')
			#$_.RemoveChild($child_node) | Out-Null
		#save our file
		#$xml.save("C:\Work\Alma\scripts\citation-xml.txt")
		#}
		
		#Write-Host "Course ID is" $course_id
	
		#$parent_xpath = 'courses'
		#$course_id = $xml.SelectSingleNode('//id')
		
		#$nodes | % {
			#$course_id = $_.SelectSingleNode('//id')
			#$course_id = $xml.SelectSingleNode('//course/id')
		#}
		
		
		# construct retrieve citation call 
		# /almaws/v1/courses/{course_id}/reading-lists/{reading_list_id}/citations/{citation_id}
		
		$retrieve_call = $url_prefix + "/almaws/v1/courses/" + $course_id + "/reading-lists/" + $list + "/citations/" + $citation

		$full_call = $retrieve_call + '?' + $queryParams
		

	
		try{
			#API call and assign result to a variable	
			[xml]$xml = Invoke-RestMethod -Method 'Get' -Uri $full_call 
			}
		catch{
			write-host "Fatal error: Get Item By Barcode" -ForegroundColor Red
			write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
			write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
			Add-Content "C:\Work\Alma\Scripts\Errors\get-errors.txt" $course
		}
		
		Write-Host $full_call
		
		#[xml]$xml = Get-Content "C:\Work\Alma\scripts\citation-xml.txt"
		
		#search for the <pages> node and remove to remove
		$parent_xpath = '//metadata'
		$nodes = $xml.SelectNodes($parent_xpath)

		$nodes | % {
			$child_node = $_.SelectSingleNode('pages')
			$_.RemoveChild($child_node) | Out-Null
		#save our file
		#$xml.save("C:\Work\Alma\scripts\citation-xml.txt")
		}

		#update citation call
		try{
			Invoke-WebRequest -Method 'PUT' -Uri $full_call -Body $xml -ContentType application/xml 
			}
		catch{
				write-host "Fatal error: Get Item By Barcode" -ForegroundColor Red
				write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
				write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red	
				Add-Content "C:\Work\Alma\Scripts\Errors\put-errors.txt" $course_id
		}
}
#close excel file
$objExcel.quit()

Write-Host "Processing complete"
