# Install-Module -Name ImportExcel
$workingPath = 'C:\GITHub\Invoices'
Set-Location $workingPath

# Import customers.xlsx
$customers = Import-Excel -Path './customers.xlsx'

# Import products.xlsx
$products = Import-Excel -Path './products.xlsx'

# Import Transactions_File.xlsx
$Transactions_File = Import-Excel -Path './Transactions_File.xlsx'

$transactionCustomers = $Transactions_File | Sort-Object 'IS_Customer' | Select-Object 'ID_Customer' -Unique

foreach ($transactionCustomer in $transactionCustomers) {
    $export = [System.Collections.ArrayList]@()
    $currentCustomer = $customers | Where-Object { $_.'ID_Customer' -eq $transactionCustomer.'ID_Customer' }
    # Write-Output "$($currentCustomer.'Billing Contact Name') of $($currentCustomer.'Business Name')"

    $currentCustomerTransactions = $Transactions_File | Where-Object { $_.'ID_Customer' -eq $transactionCustomer.'ID_Customer' }
    # Write-Output "Transaction: $($currentCustomerTransaction.'ID_Transaction')"

    foreach ($currentCustomerTransaction in $currentCustomerTransactions) {
        $currentCustomerTransactionProducts = $products | Where-Object { $_.'ID_Product' -eq $currentCustomerTransaction.'ID_Product' }
        # Write-Output "Product Name: $($currentCustomerTransactionProducts.'Product Name')"
        # Write-Output "Product Description: $($currentCustomerTransactionProducts.'Product Description')"
        # Write-Output "Product Unit Cost: $($currentCustomerTransaction.'Cost')"
        $export += [PSCustomObject]@{
            Name = $currentCustomerTransactionProducts.'Product Name'
            Description = $currentCustomerTransactionProducts.'Product Description'
            'Line Cost' = $currentCustomerTransaction.'Cost'
        }
    }
    # Dot reference template_statements.ps1
    $body = $null
    . "C:\GITHub\Invoices\template_statements.ps1"

    $body | Out-File "$env:USERPROFILE\$($transactionCustomer.'ID_Customer').html" -Force
    # $export | ConvertTo-Html | Out-File "$env:USERPROFILE\$($transactionCustomer.'ID_Customer').html"
    # Convert an HTML page into a PDF
    Start-Process "msedge.exe" -ArgumentList '--headless', "--print-to-pdf=$env:USERPROFILE\$($transactionCustomer.'ID_Customer').pdf", '--disable-extensions', '--no-pdf-header-footer', '--disable-popup-blocking', '--run-all-compositor-stages-before-draw', '--disable-checker-imaging', "$env:USERPROFILE\$($transactionCustomer.'ID_Customer').html"
    # Write-Output "Total Cost: $(($export.'Line Cost' | Measure-Object -sum).sum)"
    # Launch PDF with Edge
    Start-Process "msedge.exe" "$env:USERPROFILE\$($transactionCustomer.'ID_Customer').pdf"
}