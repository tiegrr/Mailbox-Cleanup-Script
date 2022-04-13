# Written by Bryan Phan for the IT AMER crew 8/31/21 :)
# Last updated 4/7/22

# Check if Exchange Online Module is installed
Write-Host "Checking for Exchange Online Module.. `n"
$exchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
if($exchangeModule -eq $null)
    {
        Write-Warning ("----------------------------------------------------------------------------------------------------`n`n" +
        "Exchange PowerShell Module is not installed`n" +
        "See Solution Article # 790 (https://mavenir.samanage.com/solutions/845607-how-to-connect-to-exchange-online-powershell-with-mfa) for instructions`n" +
        "Closing PowerShell script`n`n" +
        "-------------------------------------------------------------------------------------------------------------")        
        Read-Host "Press Enter to exit"
        Exit
    }
# Login to Exchange Online PS Module
Write-Host "Launching Exchange Online Module..`n-------------------------------------------------------------------------------------------------------------"
$exchangeModule
$username = Read-Host "Enter a- email (ex: a-phanb@mavenir365.onmicrosoft.com)"
Connect-ExchangeOnline -UserPrincipalName $Username

# Declaring menu function, will loop until exit option is inputted
function menu 
{
    $validAnswer = $false
        while(-not $validAnswer)
            {
                # Menu for functions
                $menuInput = Read-Host ("----------------------------------------------------------------------------------------------------`n" +
                "What are you trying to do?`n" +
                "[1] Check mailbox statistics`n" +
                "[2] Set account permissions for cleanup`n" +
                "[3] Reset account permissions back AFTER cleanup`n" +
                "[4] Start-ManagedFolderAssistant (runs email's rules and 1 day delete policies)`n" +
                "[5] Purge deleted items`n" +
                "[6] Change cleanup to different user`n" +
                "[7] Exit`n" +
                "----------------------------------------------------------------------------------------------------`n" +
                "Enter menu option")

                switch($menuInput)
                {
                "1" {
                        $validAnswer = $true
                        emailStatistics
                    }
                "2" {
                        $validAnswer = $true
                        accountPermissions

                    }
                "3" {
                        $validAnswer = $true
                        resetAccountPermissions
                    }
                "4" { 
                        $validAnswer = $true
                        startManagedFA
                    }
                "5" { 
                        $validAnswer = $true
                        clearDumpster
                    }
                "6" { 
                        $validAnswer = $true
                        changeUser
                    }
                "7" {
                        $validAnswer = $true
                        Exit
                    }
                default {Write-Warning "Enter valid menu choice`n`n"}
            }
        }
}

# Instantiating emailStatistics function - checks mailbox statistics 
function emailStatistics 
    {
        Write-Output "Showing mailbox statistics of: " $cleanupEmail
        Get-MailboxStatistics $cleanupEmail | format-list totalitemsize
        Get-MailboxStatistics $cleanupEmail | format-list totaldeleteditemsize
    }

# Instantiating accountPermissions function - sets account flags to allow for mailbox cleanup
function accountPermissions
    {
        Write-Host "Disabling Litigation Hold.."
        Set-Mailbox $cleanupEmail -LitigationHoldEnabled $false

        Write-Host "Enabling Auto-expanding Archiving.."
        Enable-Mailbox -Identity $cleanupEmail -AutoExpandingArchive

        Write-Host "Removing delay hold.."
        Set-Mailbox -Identity $cleanupEmail -RemoveDelayHoldApplied

        Write-Host "Removing email retention and deletion recovery.."
        Set-Mailbox $cleanupEmail -RetainDeletedItemsFor 0 -SingleItemRecoveryEnabled $false

        Write-Output "Finished setting account permissions for mailbox: " $cleanupEmail
    }

# Instantiating resetAccountPermissions function - resets account flags after mailbox cleanup
function resetAccountPermissions
    {
        Write-Host "Setting email retention back to 2 weeks, deletion recovery re-enabled.."   
        Set-Mailbox $cleanupEmail -RetainDeletedItemsFor 14.00:00:00 -SingleItemRecoveryEnabled $true

        Write-Host "Enabling Litigation hold.."
        Set-Mailbox $cleanupEmail -LitigationHoldEnabled $true

        Write-Output "Finished resetting account permissions for mailbox: " $cleanupEmail
    }

# Instantiating startManagedFA function - runs StartManagedFolderAssistant
function startManagedFA
    {
        Write-Host "Running Start-ManagedFolderAssistant of mailbox: " $cleanupEmail
        Start-ManagedFolderAssistant $cleanupEmail
    }

# Instantiating clearDumpster function - runs deleted items cleanup
function clearDumpster
    {
        Write-Host "Purging deleted items of mailbox: " $cleanupEmail 
        Search-Mailbox $cleanupEmail -SearchDumpsterOnly -DeleteContent -Force
    }

function changeUser
    {
        # Enter user email
        $global:cleanupEmail = Read-Host "Enter email / UPN of affected user (ex: bryan.phan@mavenir.com)"
        Write-Output $cleanupEmail
    }

# Ask for user email
$global:cleanupEmail = $null
changeUser

# Loops the menu until exit
do { menu }
until ($menuInput -eq '7')