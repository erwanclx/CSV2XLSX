Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework

$initialDirectory = [Environment]::GetFolderPath('Desktop')

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

$OpenFileDialog.InitialDirectory = $initialDirectory

$OpenFileDialog.Filter = 'Fichier CSV (*.csv)|*.csv'

$OpenFileDialog.Multiselect = $false

$response = $OpenFileDialog.ShowDialog( )

$xlsx = $OpenFileDialog.FileName.Replace(".csv",".xlsx")

if ( $response -eq 'OK' ) { Write-Host $xlsx }

Write-Host "Conversion CSV volumineux en XLSX"
$csv = $OpenFileDialog.FileName
$delimiter = ";"
$excel = New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.Refresh()
$query.Delete()
$Workbook.SaveAs($xlsx,51)
$excel.Quit()

[System.Windows.MessageBox]::Show('Le fichier converti se trouve au meme endroit que le fichier source.')