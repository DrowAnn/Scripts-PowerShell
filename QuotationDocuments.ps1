Clear-Host

#Ubicacion de los archivos
$PathLocation = "C:\Users\rober\Downloads\Cotizaciones"

#Iniciar de la aplicacion con COM
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

#Apertura del archivo de acuerdo con su tipo
$InitialFilePath = Get-ChildItem -Path "$PathLocation\*.xls"
$WorkBook = $Excel.Workbooks.Open($InitialFilePath)

#Captura del nombre de la Asesora
$WorkSheet = $WorkBook.Worksheets.Item(1)
$Asesora = $WorkSheet.Cells.Item(2,6).Value2

#Cambio de nombre a la hoja
$WorkSheet.Name = "$Asesora"

#Borrado de columnas innecesarias
$RangosABorrar = @("Q:T","J:O","H:H","E:E","C:C")
foreach ($Range in $RangosABorrar){
    $WorkSheet.Range("$Range").EntireColumn.Delete()
}

#Guardado del Archivo
$WorkBook.SaveAs("$PathLocation\$Asesora", 51) # 51 es el código para formato Excel Open XML Workbook (.xlsx)

#Cierre de los archivos
$WorkBook.Close()
$Excel.Quit()

#Liberar los objetos COM
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null

#Liberacion de las variables
Remove-Variable -Name WorkSheet
Remove-Variable -Name WorkBook
Remove-Variable -Name Excel
Remove-Variable -Name Asesora
Remove-Variable -Name PathLocation
Remove-Variable -Name InitialFilePath