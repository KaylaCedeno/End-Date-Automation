Install-Module -Name ImportExcel
Import-Module ImportExcel

Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

$ExcelLink = some_path
$Sheet = Import-Excel -Path $ExcelLink

foreach ($col in $Sheet) {

    $End = $col.'End Date'
    $Account = $col.'Work Email'

       if($Account -and $End) {
        
           try {

             Set-Mailbox $Account -CustomAttribute7 $End -ErrorAction Stop
             Write-Host "Attribute 7 updated for $Account"

           }

           catch {

             Write-Host "UNSUCESSFUL"
           }
        }
     }
