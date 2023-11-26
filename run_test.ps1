Install-Module -Name ImportExcel
$workingPath = 'C:\GITHub\Invoices'
Set-Location $workingPath

Import-Excel -Path ./