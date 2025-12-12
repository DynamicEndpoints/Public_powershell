<#
.SYNOPSIS
    Convert user mailboxes to shared mailboxes in Exchange Online.

.DESCRIPTION
    This script converts one or more user mailboxes to shared mailboxes in Exchange Online.
    When converting to a shared mailbox:
    - The mailbox type changes from UserMailbox to SharedMailbox
    - No Microsoft 365 license is required (mailbox must be under 50GB)
    - The account sign-in is automatically blocked
    - Existing permissions and data are preserved

.PARAMETER UserPrincipalName
    The UPN(s) of the user mailbox(es) to convert. Accepts multiple values.

.PARAMETER CSVPath
    Path to a CSV file containing mailboxes to convert. CSV must have a column named 'UserPrincipalName'.

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes.

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -UserPrincipalName "john.doe@contoso.com"
    Converts a single user mailbox to a shared mailbox.

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -UserPrincipalName "john.doe@contoso.com" -ResetPassword
    Converts a mailbox and resets the password for enhanced security (recommended).

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -UserPrincipalName "john.doe@contoso.com","jane.smith@contoso.com"
    Converts multiple user mailboxes to shared mailboxes.

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -CSVPath "C:\Mailboxes.csv" -ResetPassword
    Converts all mailboxes listed in the CSV file and resets passwords.

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -UserPrincipalName "john.doe@contoso.com" -WhatIf
    Shows what would happen without making actual changes.

.NOTES
    Author: Exchange Administrator
    Requirements: 
    - ExchangeOnlineManagement module
    - Microsoft.Graph.Users module (for blocking sign-in)
    - Exchange Administrator or Global Administrator role
    
    Before running:
    1. Install required modules: Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users
    2. Connect to Exchange Online: Connect-ExchangeOnline
    3. Connect to Microsoft Graph: Connect-MgGraph -Scopes "User.ReadWrite.All"
#>

[CmdletBinding(DefaultParameterSetName = 'Manual')]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Manual', ValueFromPipeline = $true)]
    [string[]]$UserPrincipalName,

    [Parameter(Mandatory = $true, ParameterSetName = 'CSV')]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CSVPath,

    [Parameter(Mandatory = $false)]
    [switch]$ResetPassword,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

begin {
    # Initialize arrays for tracking
    $successfulConversions = @()
    $failedConversions = @()
    $mailboxesToProcess = @()

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "User to Shared Mailbox Conversion Tool" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Check if connected to Exchange Online
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "[✓] Connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Host "[✗] Not connected to Exchange Online" -ForegroundColor Red
        Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
        exit
    }

    # Check if connected to Microsoft Graph
    try {
        $context = Get-MgContext -ErrorAction Stop
        if ($null -eq $context) {
            throw "No context"
        }
        Write-Host "[✓] Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        Write-Host "[✗] Not connected to Microsoft Graph" -ForegroundColor Red
        Write-Host "Please run: Connect-MgGraph -Scopes 'User.ReadWrite.All'" -ForegroundColor Yellow
        exit
    }

    if ($WhatIf) {
        Write-Host "`n[INFO] Running in WhatIf mode - no changes will be made`n" -ForegroundColor Yellow
    }

    # If CSV parameter is used, load mailboxes from CSV
    if ($PSCmdlet.ParameterSetName -eq 'CSV') {
        Write-Host "Loading mailboxes from CSV: $CSVPath" -ForegroundColor Cyan
        try {
            $csvData = Import-Csv -Path $CSVPath -ErrorAction Stop
            
            if (-not ($csvData | Get-Member -Name 'UserPrincipalName')) {
                throw "CSV must contain a 'UserPrincipalName' column"
            }
            
            $mailboxesToProcess = $csvData.UserPrincipalName
            Write-Host "[✓] Loaded $($mailboxesToProcess.Count) mailboxes from CSV`n" -ForegroundColor Green
        }
        catch {
            Write-Host "[✗] Error reading CSV: $($_.Exception.Message)" -ForegroundColor Red
            exit
        }
    }
}

process {
    # Add mailboxes from pipeline/parameter
    if ($PSCmdlet.ParameterSetName -eq 'Manual') {
        $mailboxesToProcess += $UserPrincipalName
    }
}

end {
    $totalMailboxes = $mailboxesToProcess.Count
    $counter = 0

    foreach ($upn in $mailboxesToProcess) {
        $counter++
        Write-Progress -Activity "Converting Mailboxes" -Status "Processing $counter of $totalMailboxes" -PercentComplete (($counter / $totalMailboxes) * 100)
        
        Write-Host "`n[$counter/$totalMailboxes] Processing: $upn" -ForegroundColor Cyan
        
        try {
            # Validate mailbox exists and is a user mailbox
            Write-Host "  → Checking mailbox..." -ForegroundColor Gray
            $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
            
            if ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') {
                Write-Host "  [⚠] Already a shared mailbox - skipping" -ForegroundColor Yellow
                continue
            }
            
            if ($mailbox.RecipientTypeDetails -ne 'UserMailbox') {
                Write-Host "  [✗] Not a user mailbox (Type: $($mailbox.RecipientTypeDetails)) - skipping" -ForegroundColor Red
                $failedConversions += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    Error = "Not a user mailbox (Type: $($mailbox.RecipientTypeDetails))"
                }
                continue
            }

            # Check mailbox size
            Write-Host "  → Checking mailbox size..." -ForegroundColor Gray
            $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
            
            # Extract size in bytes - handle both live and deserialized objects
            if ($stats.TotalItemSize.Value -is [string]) {
                # Parse string format like "1.234 GB (1,234,567,890 bytes)"
                if ($stats.TotalItemSize.Value -match '\(([\d,]+) bytes\)') {
                    $sizeInBytes = [long]($matches[1] -replace ',','')
                } else {
                    # Fallback: try to parse the size value directly
                    $sizeInBytes = 0
                }
            } else {
                # Try ToBytes() method if available, otherwise use ToMB() * 1MB
                try {
                    $sizeInBytes = $stats.TotalItemSize.Value.ToBytes()
                } catch {
                    # Use alternative approach with ToMB() or ToGB()
                    if ($stats.TotalItemSize.Value.PSObject.Methods['ToMB']) {
                        $sizeInBytes = $stats.TotalItemSize.Value.ToMB() * 1MB
                    } else {
                        # Last resort: parse the string representation
                        $sizeString = $stats.TotalItemSize.Value.ToString()
                        if ($sizeString -match '\(([\d,]+) bytes\)') {
                            $sizeInBytes = [long]($matches[1] -replace ',','')
                        } else {
                            $sizeInBytes = 0
                        }
                    }
                }
            }
            
            $sizeInGB = [math]::Round($sizeInBytes / 1GB, 2)
            Write-Host "  → Current size: $sizeInGB GB" -ForegroundColor Gray
            
            if ($sizeInGB -gt 50) {
                Write-Host "  [⚠] Warning: Mailbox is $sizeInGB GB (over 50GB limit for shared mailboxes)" -ForegroundColor Yellow
                Write-Host "  → Continuing conversion, but may require license to stay functional" -ForegroundColor Yellow
            }

            if ($WhatIf) {
                Write-Host "  [WHATIF] Would convert mailbox to shared" -ForegroundColor Yellow
                Write-Host "  [WHATIF] Would block sign-in for user" -ForegroundColor Yellow
                if ($ResetPassword) {
                    Write-Host "  [WHATIF] Would reset user password" -ForegroundColor Yellow
                }
            }
            else {
                # Convert mailbox to shared
                Write-Host "  → Converting to shared mailbox..." -ForegroundColor Gray
                Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                
                # Block sign-in for the account
                Write-Host "  → Blocking user sign-in..." -ForegroundColor Gray
                Update-MgUser -UserId $upn -AccountEnabled:$false -ErrorAction Stop
                
                # Reset password if requested (recommended for security)
                if ($ResetPassword) {
                    Write-Host "  → Resetting user password..." -ForegroundColor Gray
                    $newPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 16 | ForEach-Object {[char]$_})
                    $securePassword = ConvertTo-SecureString -String $newPassword -AsPlainText -Force
                    Update-MgUser -UserId $upn -PasswordProfile @{Password = $newPassword; ForceChangePasswordNextSignIn = $false} -ErrorAction Stop
                }
                
                Write-Host "  [✓] Successfully converted to shared mailbox" -ForegroundColor Green
                
                $successfulConversions += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    SizeGB = $sizeInGB
                    PasswordReset = $ResetPassword
                    ConvertedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        catch {
            Write-Host "  [✗] Error: $($_.Exception.Message)" -ForegroundColor Red
            $failedConversions += [PSCustomObject]@{
                UserPrincipalName = $upn
                Error = $_.Exception.Message
            }
        }
    }

    Write-Progress -Activity "Converting Mailboxes" -Completed

    # Summary Report
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Conversion Summary" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Total processed: $totalMailboxes" -ForegroundColor White
    Write-Host "Successful: $($successfulConversions.Count)" -ForegroundColor Green
    Write-Host "Failed: $($failedConversions.Count)" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Display successful conversions
    if ($successfulConversions.Count -gt 0) {
        Write-Host "Successfully Converted Mailboxes:" -ForegroundColor Green
        $successfulConversions | Format-Table -AutoSize
    }

    # Display failed conversions
    if ($failedConversions.Count -gt 0) {
        Write-Host "Failed Conversions:" -ForegroundColor Red
        $failedConversions | Format-Table -AutoSize
    }

    # Export results to CSV
    if (-not $WhatIf) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $exportPath = ".\ConversionResults_$timestamp.csv"
        
        $allResults = @()
        $allResults += $successfulConversions | Select-Object UserPrincipalName, SizeGB, PasswordReset, ConvertedAt, @{Name='Status';Expression={'Success'}}, @{Name='Error';Expression={''}}
        $allResults += $failedConversions | Select-Object UserPrincipalName, @{Name='SizeGB';Expression={'N/A'}}, @{Name='PasswordReset';Expression={'N/A'}}, @{Name='ConvertedAt';Expression={'N/A'}}, @{Name='Status';Expression={'Failed'}}, Error
        
        if ($allResults.Count -gt 0) {
            $allResults | Export-Csv -Path $exportPath -NoTypeInformation
            Write-Host "Results exported to: $exportPath" -ForegroundColor Cyan
        }
    }

    Write-Host "`nNext Steps:" -ForegroundColor Yellow
    Write-Host "1. Remove Microsoft 365 licenses from converted accounts (if mailbox < 50GB)" -ForegroundColor White
    if (-not $ResetPassword) {
        Write-Host "2. SECURITY: Consider resetting passwords - original credentials still work!" -ForegroundColor Red
        Write-Host "   Run again with -ResetPassword switch for enhanced security" -ForegroundColor Yellow
        Write-Host "3. Grant Full Access permissions to users who need access" -ForegroundColor White
    }
    else {
        Write-Host "2. Grant Full Access permissions to users who need access" -ForegroundColor White
    }
    Write-Host "   Example: Add-MailboxPermission -Identity shared@contoso.com -User user@contoso.com -AccessRights FullAccess" -ForegroundColor Gray
    Write-Host "3. Verify users can access the shared mailbox in Outlook`n" -ForegroundColor White
}
