##This is a Script to take specific quotation documents from a single consultant an convert them to the working format

Clear-Host

#Files location
$PathLocation = "C:\Users\rober\Downloads\Cotizaciones"

#Start applications with COM
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

#Open files by file extension
$InitialFilePath = Get-ChildItem -Path "$PathLocation\*.xls"
$WorkBook = $Excel.Workbooks.Open($InitialFilePath)

#Get the consultant´s name
$WorkSheet = $WorkBook.Worksheets.Item(1)
$Consultant = $WorkSheet.Cells.Item(3,6).Value2

#Change the name of the sheet
$WorkSheet.Name = "$Consultant"

#Delete unnecesary columns
$DeleteRange = @("Q:T","J:O","H:H","E:E","C:C")
foreach ($Range in $DeleteRange){
    $WorkSheet.Range("$Range").EntireColumn.Delete()
}

#Delete dots of price numbers
$FindText = ".*"
$ReplaceText = ""
$WorkSheet.Range("C:C").Replace($FindText, $ReplaceText)

#Formatting texts
$WorkSheet.Range("C:C").NumberFormat = "$#.##0"
$WorkSheet.Range("E:E").NumberFormat = "dd/mm/yyyy"
$WorkSheet.Range("F:F").NumberFormat = "0"

#Get Table Range
$ColumnRange = $worksheet.Columns.Item("A")
$EndRow = $ColumnRange.Cells($ColumnRange.Rows.Count, 1).End(-4162).Row # -4162 its the code for search from the bottom to the top
$TableRange = $WorkSheet.Range("A1:G$EndRow")

#Create table
$Table = $WorkSheet.ListObjects.Add(1, $TableRange, $null, 1) # The first 1 declares the table object type and the last 1 set the true for headers
$Table.Name = "Tabla1"

#Save file
$WorkBook.SaveAs("$PathLocation\$Consultant", 51) # 51 its the code for file extension, Excel Open XML Workbook (.xlsx)

#Close files
$WorkBook.Close()
$Excel.Quit()

#Release COM Objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ColumnRange) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Table) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null

#Release variables
Remove-Variable -Name ColumnRange
Remove-Variable -Name Table
Remove-Variable -Name WorkSheet
Remove-Variable -Name WorkBook
Remove-Variable -Name Excel
Remove-Variable -Name Consultant
Remove-Variable -Name PathLocation
Remove-Variable -Name InitialFilePath
Remove-Variable -Name DeleteRange
Remove-Variable -Name FindText
Remove-Variable -Name ReplaceText
Remove-Variable -Name EndRow
Remove-Variable -Name TableRange