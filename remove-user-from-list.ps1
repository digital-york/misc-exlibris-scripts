# PEH Oct 2018
# PS Script to remove user fromreading lists
# 
#course_id			xs:long	The identifier of the Course.
#reading_list_id	xs:long	The identifier of the Reading List.
#primary_id			xs:string	The primary identifier of the user.
# 

#load the System Web Assembly - required for encoding action below
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

#define variables
$url_prefix = "https://api-na.hosted.exlibrisgroup.com/"
$queryParams = [System.Web.HttpUtility]::UrlEncode('apikey') + '=' + [System.Web.HttpUtility]::UrlEncode('l7xxa15adcbc13f9415b895752a24a91f827');
#$api_key = [System.Web.HttpUtility]::UrlEncode($key) 
$file = "C:\Work\Alma\scripts\list-owners\cm541 lists.xlsx"
$sheetName = "readingListList"
$owner = "cm541"

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


#loop through sheet and act on each row
	for ($i=1; $i -le $rowMax-1; $i++)
	{
		$course = $sheet.Cells.Item($rowCourse+$i,$colCourse).text
		$list= $sheet.Cells.Item($rowList+$i,$colList).text
	
		#test for empty course code
		if (!$course){
			Write-Host "Null course. Exiting."
			Exit
		}
	
		#Write-Host "Processing row number" ($rowCourse+$i) "course code " $course
	
		#course id is required to update citation
		#this is retrieved via a courses api search using the module code
		$course_id_call = $url_prefix + "almaws/v1/courses?q=code~" + $course + '&' + $queryParams
		
		Write-Host $course_id_call
			
		try
		{
			#get returned xml
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $course_id_call 
		}
		catch
		{
			Write-Host "Fatal error: Retrieve Course " -ForegroundColor Red
		}
		
		#retrieve course id from xml
		$course_id = $xml.courses.course.id
		
		#retrieve list id
		$list_id_call = $url_prefix + "almaws/v1/courses/" + $course_id + "/reading-lists/" + '?' +$queryParams
		
		Write-Host $list_id_call
		
		try
		{
			#get returned xml
			[xml]$xml = Invoke-RestMethod -Method 'GET' -Uri $list_id_call
		}
		catch
		{
			Write-Host "Fatal error: Retrieve List " -ForegroundColor Red
		}
		
		#count number of lists for current course
		$num_lists = $xml.CreateNavigator().Evaluate('count(//id)')
				
		#if one list, go ahead and retrieve the id
		if ($num_lists -eq 1)
			{
				Write-Host "Course " $course "has a single list"
				$list_id = $xml.reading_lists.reading_list.id
				Write-Host $list_id
				#now we have the course and list id, we can perform the delete owner call
				$delete_call = $url_prefix + "almaws/v1/courses/" + $course_id + "/reading-lists/" + $list_id + "/owners/" + $owner + "?" + $queryParams
				try
					{
						Write-Host $delete_call
						Invoke-RestMethod -Method 'DELETE' -Uri $delete_call
					}
				catch
					{
						Write-Host "Fatal error: Delete Owner " -ForegroundColor Red
						write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
						write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
					}	
			}
		else
			{
				Write-Host $num_lists "lists"
				#get list id for each list
				for ($j=0; $j -le $num_lists-1; $j++)
				{
					$list_code = $xml.reading_lists.reading_list.code[$j]
					$list_id = $xml.reading_lists.reading_list.id[$j]
					Write-Host "list id " $j "is" $list_id " for list code" $list_code
					
					#now we have the course and list id, we can perform the delete owner call
					$delete_call = $url_prefix + "almaws/v1/courses/" + $course_id + "/reading-lists/" + $list_id + "/owners/" + $owner + "?" + $queryParams
					
					Write-Host "delete call " $j "is " $delete_call 
					
					try
					{
						Invoke-RestMethod -Method 'DELETE' -Uri $delete_call
					}
					catch
					{
						Write-Host "Fatal error: Delete Owner " -ForegroundColor Red
						write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
						write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
					}	
				}
			}
		}

#close excel file
$objExcel.quit()

Write-Host "Processing Complete"



