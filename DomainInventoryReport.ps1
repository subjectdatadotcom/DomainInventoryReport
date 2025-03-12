<#
.SYNOPSIS
This PowerShell script inventories objectes for specified domains in Exchange Online, providing comprehensive insights without making any changes to the configurations.

.DESCRIPTION 
The script ensures that necessary PowerShell modules (ExchangeOnlineManagement and MSOnline) are installed and loaded, connects to Exchange Online, and imports domains from a CSV file for detailed review. It gathers email properties such as UserPrincipalName and primary SMTP addresses, cataloguing these for various recipient types like UserMailbox, MailUser, and GroupMailbox. The script aims to log all findings to aid in administrative reviews and support compliance, highlighting potential areas for future interventions.

.NOTES 
The script requires administrative credentials for Exchange Online and is designed for administrators who need to audit and maintain accurate records of recipient configurations across Exchange Online.

.AUTHOR 
SubjectData

.EXAMPLE 
.\DomainInventoryReport.ps1 Executes the script, compiling email address configurations according to the domains listed in 'Domains.csv', generating a detailed report. This report includes current settings and identifies all processed entities without making any modifications.
#>

# Ensure ExchangeOnlineManagement module is installed and imported
$exchangeModule = "ExchangeOnlineManagement"

if (-not (Get-Module -Name $exchangeModule -ListAvailable)) {
    Write-Host "$exchangeModule module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name $exchangeModule -Force -Scope CurrentUser
}

Import-Module $exchangeModule -Force
Write-Host "$exchangeModule module successfully loaded." -ForegroundColor Green

# Ensure MSOnline module is installed and imported
$msolModule = "MSOnline"

if (-not (Get-Module -Name $msolModule -ListAvailable)) {
    Write-Host "$msolModule module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name $msolModule -Force -Scope CurrentUser
}

Import-Module $msolModule -Force
Write-Host "$msolModule module successfully loaded." -ForegroundColor Green

# Connect to Exchange Online
try {
    Connect-ExchangeOnline 
    Connect-MsolService
} catch {
    Write-Host "Failed to connect to Exchange Online. Please check your credentials and try again." -ForegroundColor Red
    exit
}

# Get the directory of the current script
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define the location of the CSV file containing OneDrive user emails
$XLloc = "$myDir\"

try {
    # Import the list of OneDrive users from the CSV file
    $Domains = import-csv ($XLloc + "Domains.csv").ToString() | Select-Object -ExpandProperty Domain
} catch {
    # Handle the error if the CSV file is not found
    Write-Host "No CSV file to read" -BackgroundColor Black -ForegroundColor Red
    exit
}

# Initialize report array
$report = @()

foreach ($DomainToReplace in $Domains) {
    $DomainToReplace2 = "*@" + $DomainToReplace
    Write-Host "Processing domain: $DomainToReplace" -ForegroundColor Cyan

    # Get recipients with matching email domain (licensed accounts with mailboxes, groups, etc)
    $RecipientList = Get-EXORecipient -ResultSize unlimited | Where-Object { $_.EmailAddresses -like $DomainToReplace2 }

    $report = @()

    foreach ($Recipient in $RecipientList) {
        $GUID = $Recipient.ExternalDirectoryObjectID.ToString()
        $PrimarySMTP = $Recipient.PrimarySMTPAddress
        $DisplayName = $Recipient.DisplayName
        $RecipientType = $Recipient.RecipientType
        $RecipientDetails = $Recipient.RecipientTypeDetails

        $UserObj = New-Object PSObject
        $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
        $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
        $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $PrimarySMTP
        $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
        $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
        $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                
        # Handling UserMailbox - Completed - Green
        if ($RecipientDetails -eq "UserMailbox") {
            $ThisMailbox = Get-Mailbox -Identity $GUID
            $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced

            # Initialize Warning and Error logs as empty arrays
            $WarningLog = @()
            $ErrorLog = @()

             Write-Host "User Mailbox - checking " $ThisMailUser.PrimarySMTPAddress -ForegroundColor Green

            try {
                Write-Host "Checking User Mailbox addresses on " $ThisMailUser.DisplayName
                foreach ($Address in $ThisMailbox.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced
                    $report += $UserObj
                }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                Write-Host $ThisMailbox.UserPrincipalName -ForegroundColor DarkMagenta
                # Capture General Errors
                $ErrorLog += "General Error: $($_.Exception.Message)"                
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling MailUser - Completed - Yellow
        elseif ($RecipientType -eq "MailUser") {
             try {
                $ThisMailUser = Get-MailUser -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailUser.IsDirSynced

                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()

                Write-Host "Mail user - checking " $ThisMailUser.DisplayName -ForegroundColor Yellow
                foreach ($Address in $ThisMailUser.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced
                    $report += $UserObj
                }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General MailUser Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling Security Groups - Completed - Magenta
        elseif (($RecipientType -match "MailUniversalDistributionGroup|MailUniversalSecurityGroup") -and ($RecipientDetails -match "MailUniversalDistributionGroup|MailUniversalSecurityGroup")) {
            try {
                # Fetch Distribution/Security Group Information
                $ThisDistroGroup = Get-DistributionGroup -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisDistroGroup.IsDirSynced

                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()

                Write-Host "MailUniversalDistributionGroup / MailUniversalSecurityGroup User addresses on " $ThisDistroGroup.DisplayName -ForegroundColor Magenta

                foreach ($Address in $ThisDistroGroup.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced
                    $report += $UserObj
                }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    Write-Host "Entered in logs"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }


        # Handling MailContact - Completed - Dark Green
        elseif ($RecipientType -eq "MailContact") {
            try {
                # Fetch Mail Contact Information
                Write-Host "Mail Contact found - logging ONLY - " $Recipient.Identity -ForegroundColor DarkGreen
                $ThisMailContact = Get-MailContact -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailContact.IsDirSynced
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        <#
                $ThisPrimarySMTP = $ThisMailContact.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
       
                # -- Primary SMTP Logging (No Change) --
                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    Write-Host "Current Primary SMTP is " $ThisPrimarySMTP
                    $ActionValue = "Primary SMTP Found for Mail Contact - Logging Only - " + $ThisPrimarySMTP
                    Write-Host $ActionValue
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value $ActionValue
                }
        #>
                foreach ($Address in $ThisMailContact.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced
                    $report += $UserObj
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
        
            }
            catch {
                $ErrorLog += "General Mail Contact Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }
        
        # Handling DynamicDistributionGroup - Completed - White
        elseif ($RecipientType -eq "DynamicDistributionGroup") {
            try {
                # Fetch Dynamic Distribution Group Information
                $ThisDistroGroup = Get-DynamicDistributionGroup -Identity $GUID
                Write-Host "Checking Dynamic Distribution Group - " $ThisDistroGroup.DisplayName -ForegroundColor White
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        
                foreach ($Address in $ThisDistroGroup.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $report += $UserObj
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General Dynamic Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling Group Mailboxes - Completed
        elseif (($RecipientType -eq "MailUniversalDistributionGroup") -and ($RecipientDetails -eq "GroupMailbox")) {
            try {
                # Fetch Group Mailbox (Unified Group) Information
                $ThisUnifiedGroup = Get-UnifiedGroup -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisUnifiedGroup.IsDirSynced
                Write-Host "Checking Office 365 Group - " $ThisUnifiedGroup.DisplayName
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        
                foreach ($Address in $ThisUnifiedGroup.EmailAddresses) {
                    $UserObj = New-Object PSObject
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
                    $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
                    $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $Address
                    $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
                    $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                    $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced
                    $report += $UserObj
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
        
            }
            catch {
                $ErrorLog += "General Office 365 Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }
        else {
            <# Action when all if and elseif conditions are false #>
			$ActionValue = "************Unrecognized recipient type!"
			Write-Host $ActionValue -foregroundcolor DarkRed
			$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue
        }

        $report += $UserObj       
    }

    #Let's log any remaining user accounts that are using the domain in some way, but don't have any Exchange attributes, and are therefore not Recipients
    Write-Host "Let's check any user accounts that are using the domain but have no mail attributes" -foregroundcolor cyan
    $UsersWithDomain = Get-MsolUser -All -DomainName $DomaintoReplace
    ForEach ($User in $UsersWithDomain) {
	    #Is User not a recipient?
	    $UserCheck = Get-Recipient $User.UserPrincipalName -ErrorAction SilentlyContinue
	    If ($UserCheck -eq $NULL) {
		    Write-Host "User found with domain - " $User.DisplayName
		    $ThisUPN = $User.UserPrincipalName
		    $UserObj = New-Object PSObject
		    $UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
		    $UserObj | Add-Member -membertype NoteProperty -Name "UPN" -Value $User.UserPrincipalName
		    $UserObj | Add-Member -membertype NoteProperty -Name "DisplayName" -Value $User.DisplayName
		    $UserObj | Add-Member -membertype NoteProperty -Name "RecipientType" -Value "User ONLY - Manual intervention required"
		    $UserObJ | Add-Member -membertype NoteProperty -Name "Action" -Value "***Manual intervention required for this object!"
		    $UserObj | Add-Member -Membertype NoteProperty -Name "IsObjectSyncedAADC" -Value $User.LastDirSyncTime
		    $report += $UserObj	
	    }
    }
    Write-Host "Finished checking all remaining users with the domain" -foregroundcolor Cyan

    # Sort and select unique entries based on 'PrimaryEmailAddress'
    $uniqueReport = $report | Sort-Object -Property PrimaryEmailAddress | Select-Object -Unique -Property PrimaryEmailAddress, GUID, DisplayName, RecipientType, RecipientTypeDetails, IsObjectSyncedAADC, Time
    # Export the unique results to a CSV file
    $domainNameForFile = $DomainToReplace -replace "[^a-zA-Z0-9]", "_"  # Replace non-alphanumeric characters to ensure valid file name
    $FileName = $XLloc+${domainNameForFile}+"_Report-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv"
    $uniqueReport | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8

    Write-Host "${domainNameForFile} report exported"
    #read-host
}

Write-Host "Script execution completed. Log file: $FileName" -ForegroundColor Green

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
