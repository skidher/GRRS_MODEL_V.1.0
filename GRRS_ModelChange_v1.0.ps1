Clear
###################################################################
# Load Exchange Snapin modules for testing purposes in ISE
###################################################################

# Clear-Variable "exist" -Scope local
# Call Exchange Snapin
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://INHEXCH04.eu.boehringer.com/PowerShell/ -Authentication Kerberos
#import-pssession $session -AllowClobber
#add-pssnapin microsoft.exchange*
#test-webservicesconnectivity -clientaccessserver INHEXCH04 –trustanysslcertificate

###################################################################
# Ask for mailbox to change and check existence and current model
###################################################################

function Check-Grrs {
$global:model = $null
$grrsmb = Read-Host -Prompt "Enter the GRRS mailbox to be changed" -erroraction SilentlyContinue
$exist = [bool](Get-mailbox $grrsmb -erroraction SilentlyContinue)
while($exist -ne $True)
{
    if ($exist -eq $False) {Write-Host "Mailbox" $grrsmb "exists!"}
    $grrsmb = Read-Host -Prompt "Enter the GRRS mailbox to be changed"
    $exist = [bool](Get-mailbox $grrsmb -erroraction SilentlyContinue)
}    

$chkmailbox = (Get-CalendarProcessing $grrsmb | select ForwardRequestsToDelegates,AllBookInPolicy,AllRequestInPolicy,ResourceDelegates,BookInPolicy) 
#Get-CalendarProcessing $grrsmb | select ForwardRequestsToDelegates,AllBookInPolicy,AllRequestInPolicy,ResourceDelegates,BookInPolicy # proof of results

Write-Host "`n==============="
Write-Host "Mailbox exists!"
Write-Host "===============`n"

If ($chkmailbox.ForwardRequestsToDelegates -eq $false -and $chkmailbox.AllBookInPolicy -eq $true  -and $chkmailbox.AllRequestInPolicy -eq $false) {
$global:model = "OPEN"
Write-Host "This GRRS Mailbox is $model model`n" -ForegroundColor Black -BackgroundColor Yellow
} elseif ($chkmailbox.ForwardRequestsToDelegates -eq $true -and $chkmailbox.AllBookInPolicy -eq $false  -and $chkmailbox.AllRequestInPolicy -eq $true) {
$global:model = "MODERATED"
write-host "This GRRS Mailbox is $model model`n" -ForegroundColor Red -BackgroundColor Yellow
} elseif ($chkmailbox.ForwardRequestsToDelegates -eq $true -and $chkmailbox.AllBookInPolicy -eq $false -and $chkmailbox.AllRequestInPolicy -eq $false) {
$global:model = "CLOSED"
write-host "This GRRS Mailbox is $model model`n" -ForegroundColor Red -BackgroundColor Yellow
} else {write-host "GRRS Mailbox does not match any model caracteristics!" -ForegroundColor Red}


$confirmation = Read-Host "Are you sure you want to proceed? [y/n]"
while($confirmation -ne "y" -and $confirmation -ne $null)
{
    if ($confirmation -eq 'n') {Write-Host "CANCELED! No changes were made"}
    $confirmation = Read-Host "Are you sure you want to proceed? [y/n]"
}
} # end Check-grrs
Check-Grrs

# Start of the menu options
function Show-Menu
{
     param (
           [string]$Title = 'GRRS Mailbox Options Menu'
     )
     cls
     Write-Host "================ $Title ================`n"
     if ($global:model -eq 'OPEN'){
     Write-Host "M: Press 'M' to change to MODERATED model."
     Write-Host "C: Press 'C' to change to CLOSED model."
     Write-Host "Q: Press 'Q' to quit.`n"
     } elseif ($global:model -eq 'MODERATED'){
     Write-Host "O: Press 'O' to change to OPEN model."
     Write-Host "C: Press 'C' to change to CLOSED model."
     Write-Host "Q: Press 'Q' to quit.`n"
     } elseif ($global:model -eq 'CLOSED') {
     Write-Host "O: Press 'O' to change to OPEN model."
     Write-Host "M: Press 'M' to change to MODERATED model."
     Write-Host "Q: Press 'Q' to quit.`n"
     } else {
     write-host "GRRS Mailbox does not match any model caracteristics!`n" -ForegroundColor Red
     Write-Host "Q: Press 'Q' to quit.`n"
     }
}

do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
             'O' {
                cls
                'You chose OPEN'
           } 'M' {
                cls
                'You chose MODERATED'
           } 'C' {
                cls
                'You chose CLOSED'
           } 'q' {
                return
           }
     }
     pause
}
until ($input -eq 'q')

###################################################################
# Start of functions to change models
###################################################################

function changeToOpen {}


function changeToModerated {}


function changeToClosed {}