$file = "C:\Work\Alma-D\ids_full.xlsx"
$sheetName = "Sheet1"
$xmlFileName = "C:\Work\Alma-D\delete-dlib-recs.xml"


#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

#create XML file
#$xmlDoc = [System.Xml.XmlDocument](Get-Content $xmlFileName);
#[xml]$xmlDoc = New-Object system.Xml.XmlDocument
#$xmlDoc.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?><ListRecords></ListRecords>")

#Count max row
$rowMax = ($sheet.UsedRange.Rows).count

Write-Host "We have " ($rowMax) " rows to process"


#Declare the starting position
$rowStart, $startRow = 0,1

$data = ''

for ($i=1; $i -le $rowMax; $i++){


    Write-Host "Processing row " $i


    $cur_id = $sheet.Cells.Item($rowStart+$i,$startRow).text
    Write-Host $cur_id

    $data = $data +'<record><header status="deleted"><identifier>' + $cur_id + '</identifier><datestamp>2021-06-14</datestamp></header><metadata></metadata></record>'
            
@"
<?xml version='1.0' encoding='UTF-8'?>
<?xml-stylesheet type='text/xsl' href='/oai2.xsl'?>
<OAI-PMH xmlns="http://www.openarchives.org/OAI/2.0/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.openarchives.org/OAI/2.0/ http://www.openarchives.org/OAI/2.0/OAI-PMH.xsd">
<ListRecords xmlns="">
$data
</ListRecords>
</OAI-PMH>
"@ | Out-File $xmlFileName

}

#close excel file
$objExcel.quit()

Write-Host "Processing Complete"




