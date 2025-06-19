# PEH Aug 2024
# PS script to delete old courses/lists
# uses course API 

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-eu.hosted.exlibrisgroup.com"
$queryParams = [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('l8xx8b049ec8d68b47a1b130d368e367b4a0');
$file = "C:\Work\inactive-no-list.xlsx"
$sheetName = "courses"

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

#loop to get values and store
#retrieve ID associated with course code
	for ($i=1; $i -le $rowMax-1; $i++)
	{
		$course = $sheet.Cells.Item($rowCourse+$i,$colCourse).text

        Write-Host "Processing row" $i 
		
		#test for empty course code
		if (!$course){
			Write-Host "Null course. Exiting."
			Exit
		}
			
		#course id is required to update citation
		#this is retrieved via a courses api search using the module code
		$course_id_call = $url_prefix + "/almaws/v1/courses?q=code~" + $course + '&' + $queryParams

        Write-Host $course_id_call
		
		#get returned xml
		[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $course_id_call 
		
		#retrieve values from xml response
        $course_id = $xml.courses.course.id
        $course_name = $xml.courses.course.name
        $dept = $xml.courses.course.academic_department.desc
		$course_date = [DateTime] $xml.courses.course.created_date

		$del_date = [DateTime] '2019-01-01'

		
		if ($course_date -lt $del_date) {
			Write-Host "Date 1 is less than Date 2"
            $course + ',' + $course_name +',' + $dept +',' + $course_date | Out-File -Encoding utf8 "C:\Work\Data.csv" -Append
        
		}


}
#close excel file
$objExcel.quit()

Write-Host "Processing complete"
