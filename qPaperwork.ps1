<#
.SYNOPSIS
    qPaperwork

.DESCRIPTION
    Prints relevant paperwork from list of order numbers

.NOTES
    Author: Joshua Earl
    Version 1.0
#>

function showInfoHeader
{
    Write-Host "QPaperwork`n" -ForegroundColor yellow -BackgroundColor blue
    Write-Host "Type in an Order ID to add it to the list"
    Write-Host "Type 'remove' to remove an ID"
    Write-Host "Type 'done' when finished`n"
}

$basedir = "//networkshare/ORDERS"
$docs = @('customerDetails.doc','install_sheet.pdf','checklist.docx')
$orderIDList = @()
$soffice = "C:\Program Files (x86)\OpenOffice 4\program\soffice.exe"
$switches = "-p"

clear
showInfoHeader
Do
{
    $userInput = Read-Host 'Input'
    clear
    showInfoHeader
    if ($userInput.length -eq 6 -and $userInput -match '^\d+$') {
        if ($orderIDList.Contains($userInput)) {
            Write-Host "Error - Duplicate Order ID entered: $userInput`n"
        }
        else {
            Write-Host "Added Order ID: $userInput`n"
            $orderIDList += $userInput
        }
    }
    elseif ($userInput -eq 'remove') {
            Write-Host "What Order ID needs be removed?`n"
            Write-Output $orderIDList
            Write-Output ""
            $removeOrderID = Read-Host 'Enter ID to be removed'
            clear
            showInfoHeader
            $tempvar = $orderIDList -ne "$removeOrderID"
            $orderIDList = $tempvar
            Write-Host "Removed Order ID: $removeOrderID`n"
    }
    elseif (($userInput -ne 'done') -and ($userInput -ne 'remove')) {
            Write-Host "Error - Please enter a 6 digit order ID!`n"
    }
    if ($orderIDList -ne '0') {
        Write-Output $orderIDList
        Write-Output ""
    }
} Until ($userInput -eq 'done')

clear
foreach ($orderID in $orderIDList) {
    Write-Host "Order #$orderID" -ForegroundColor yellow -BackgroundColor blue
    foreach ($doc in $docs) {
        $path = "$basedir/$orderID/$doc"
        if (Test-Path $path -PathType Leaf) {
            Write-Host "Printing: $doc" -ForegroundColor green
            & $soffice $switches $path
        }
        else {
            Write-Warning "Can't find Order #$orderID $doc"
        }
        Start-Sleep -s 1.5
    }
}