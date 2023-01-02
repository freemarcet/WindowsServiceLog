## Enter service name here
$ServiceName = 'MSSQLSERVER'

## Get Current Timestamp
$CurrentTime = Get-Date

## Enter path to create/write to Excel file
## $Path = 'C:\Users\
## $PathForNewFile = 'C:\Users


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $True
$workbook = $excel.Workbooks.add() 
$sheet = $workbook.worksheets.Item(1) 
$workbook.WorkSheets.item(1).Name = "LogPage"
$sheet = $workbook.WorkSheets.Item("LogPage")
$sheet.Cells.Item(1,1) = 'Status'
$sheet.Cells.Item(1,2) = 'Timestamp'

# Find the last used cell
$count = 1
if($sheet.Cells.Item($count,1) = null)
{
    $nextRow = $count
    #do stuff on this row   
}
else
{
    $count++
}


## Store all service information
$ServiceInfo = Get-Service -Name $ServiceName

## Only show status from $ServiceInfo
$ServiceInfo.Status

## If the server is not running (ne)
if ($ServiceInfo.Status -ne 'Running') {

	## Option to automate restarting the service (uncomment line below)
	## Start-Service -Name $ServiceName
    $sheet.Cells.Item($nextRow,1) = $ServiceInfo.Status
    $sheet.Cells.Item($nextRow,2) = $CurrentTime
    
    

} else { ## If the Status is anything but Running
	## Write to the console the service is already running
	Write-Host 'The service is already running.'
}
