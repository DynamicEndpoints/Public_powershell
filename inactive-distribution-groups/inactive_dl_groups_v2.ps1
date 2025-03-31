param(
    [int]$InactiveDays = 90,
    [int]$MessageTraceDays = 10,
    [string]$ReportPath = [Environment]::GetFolderPath("MyDocuments"),
    [string[]]$GroupFilter,
    [string[]]$DomainFilter
)

# Initialize logging
$logFile = Join-Path -Path $ReportPath -ChildPath "InactiveDLReport_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
Start-Transcript -Path $logFile -Append

try {
    # Connect to Exchange Online
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green

    # Set time thresholds
    $inactiveDate = (Get-Date).AddDays(-$InactiveDays)
    $startDate = (Get-Date).AddDays(-$MessageTraceDays)
    $endDate = Get-Date

    Write-Host "Using inactivity threshold: $InactiveDays days" -ForegroundColor Cyan
    Write-Host "Using message trace window: $MessageTraceDays days" -ForegroundColor Cyan

    # Get distribution groups with optional filtering
    Write-Host "Retrieving distribution groups..." -ForegroundColor Cyan
    $groupFilterParams = @{
        ResultSize = 'Unlimited'
    }

    if ($GroupFilter) {
        $groupFilterParams['Identity'] = $GroupFilter
    }

    $groups = Get-DistributionGroup @groupFilterParams
    
    # Apply domain filtering if specified
    if ($DomainFilter) {
        $groups = $groups | Where-Object {
            $groupDomain = ($_.PrimarySmtpAddress -split '@')[1]
            $DomainFilter -contains $groupDomain
        }
    }

    $totalGroups = $groups.Count
    Write-Host "Found $totalGroups distribution groups to process" -ForegroundColor Green

    # Create array to store results
    $inactiveGroups = @()
    $reportDate = Get-Date -Format "yyyy-MM-dd_HHmmss"

    # Initialize progress bar variables
    $progressCounter = 0
    $progressParams = @{
        Activity = "Scanning Distribution Groups"
        Status = "Processing groups..."
        PercentComplete = 0
    }

    foreach ($group in $groups) {
        # Update progress bar
        $progressCounter++
        $progressParams.PercentComplete = ($progressCounter / $totalGroups * 100)
        $progressParams.Status = "Processing group $progressCounter of $totalGroups : $($group.DisplayName)"
        Write-Progress @progressParams

        # Get group stats
        $stats = Get-MailboxFolderStatistics -Identity $group.PrimarySmtpAddress -FolderScope Inbox -ErrorAction SilentlyContinue
        
        # Get DL Members and Owners
        try {
            $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress
            $memberCount = $members.Count
            $memberDetails = $members | ForEach-Object {
                "$($_.DisplayName) ($($_.PrimarySmtpAddress))"
            }

            $owners = Get-DistributionGroup -Identity $group.Identity | Select-Object -ExpandProperty ManagedBy
            $ownerDetails = foreach ($owner in $owners) {
                try {
                    $ownerObject = Get-Recipient $owner -ErrorAction Stop
                    "$($ownerObject.DisplayName) ($($ownerObject.PrimarySmtpAddress))"
                }
                catch {
                    "$owner (Unable to resolve details)"
                }
            }
        }
        catch {
            Write-Warning "Could not get member/owner details for group: $($group.DisplayName). Error: $($_.Exception.Message)"
            $memberCount = "Error getting count"
            $memberDetails = @("Error retrieving members")
            $ownerDetails = @("Error retrieving owners")
        }

        # Get last sent and received emails (last 10 days)
        try {
            $lastReceived = Get-MessageTrace -RecipientAddress $group.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate |
                Sort-Object Received -Descending | Select-Object -First 1

            $lastSent = Get-MessageTrace -SenderAddress $group.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate |
                Sort-Object Received -Descending | Select-Object -First 1

            if ($lastReceived) {
                $lastReceivedDetail = Get-MessageTraceDetail -MessageTraceId $lastReceived.MessageTraceId -RecipientAddress $lastReceived.RecipientAddress -ErrorAction SilentlyContinue
            }
            if ($lastSent) {
                $lastSentDetail = Get-MessageTraceDetail -MessageTraceId $lastSent.MessageTraceId -RecipientAddress $lastSent.RecipientAddress -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Warning "Could not get message trace for group: $($group.DisplayName). Error: $($_.Exception.Message)"
            $lastReceived = $null
            $lastSent = $null
        }

        # If no recent activity or no stats available, consider it potentially inactive
        if (($stats.LastModifiedTime -eq $null) -or ($stats.LastModifiedTime -lt $inactiveDate)) {
            try {
                # Get additional group details
                $groupDetails = Get-DistributionGroup -Identity $group.Identity
                $emailAddresses = ($groupDetails.EmailAddresses | Where-Object {$_ -match "smtp:"}) -join "; "
            }
            catch {
                Write-Warning "Could not get additional details for group: $($group.DisplayName)"
            }
            
            $inactiveGroups += [PSCustomObject]@{
                Name = $group.DisplayName
                EmailAddress = $group.PrimarySmtpAddress
                AllEmailAddresses = $emailAddresses
                MemberCount = $memberCount
                Members = ($memberDetails -join "`n")
                Owners = ($ownerDetails -join "`n")
                LastModified = $group.WhenChanged
                Created = $group.WhenCreated
                ManagedBy = ($group.ManagedBy -join ';')
                LastActivityDate = $stats.LastModifiedTime
                LastEmailReceived = if ($lastReceived) { $lastReceived.Received } else { "No emails received in last $MessageTraceDays days" }
                LastEmailReceivedFrom = if ($lastReceived) { $lastReceived.SenderAddress } else { "N/A" }
                LastEmailReceivedSubject = if ($lastReceived) { $lastReceived.Subject } else { "N/A" }
                LastEmailSent = if ($lastSent) { $lastSent.Received } else { "No emails sent in last $MessageTraceDays days" }
                LastEmailSentTo = if ($lastSent) { $lastSent.RecipientAddress } else { "N/A" }
                LastEmailSentSubject = if ($lastSent) { $lastSent.Subject } else { "N/A" }
                LastReceivedStatus = if ($lastReceivedDetail) { ($lastReceivedDetail.Event -join '; ') } else { "N/A" }
                LastSentStatus = if ($lastSentDetail) { ($lastSentDetail.Event -join '; ') } else { "N/A" }
                HiddenFromGAL = $group.HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
                AcceptMessagesOnlyFrom = ($group.AcceptMessagesOnlyFrom -join ';')
                AcceptMessagesOnlyFromDLMembers = ($group.AcceptMessagesOnlyFromDLMembers -join ';')
                CustomAttribute1 = $group.CustomAttribute1
                CustomAttribute2 = $group.CustomAttribute2
                Notes = $group.Notes
            }
        }
    }

    # Clear progress bar
    Write-Progress -Activity "Scanning Distribution Groups" -Completed

    # Create reports subfolder
    $reportsFolder = Join-Path -Path $ReportPath -ChildPath "DistributionGroupReports"
    if (-not (Test-Path -Path $reportsFolder)) {
        New-Item -Path $reportsFolder -ItemType Directory
    }

    # Export results to CSV
    $csvPath = Join-Path -Path $reportsFolder -ChildPath "InactiveDistributionGroups_$reportDate.csv"
    $inactiveGroups | Export-CSV -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV report saved to: $csvPath" -ForegroundColor Green

    # Create detailed HTML report
    $htmlReport = Join-Path -Path $reportsFolder -ChildPath "InactiveDistributionGroups_$reportDate.html"
    $htmlContent = @"
<html>
<head>
    <title>Inactive Distribution Groups Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { margin-bottom: 20px; }
        .highlight { background-color: #fff3cd; }
        .members-list { max-height: 200px; overflow-y: auto; }
        .section { margin-bottom: 30px; }
        .group-details { margin-bottom: 40px; padding: 10px; border: 1px solid #ddd; }
    </style>
</head>
<body>
    <h1>Inactive Distribution Groups Report</h1>
    <div class="summary">
        <h2>Summary</h2>
        <p>Report Generated: $(Get-Date)</p>
        <p>Total Groups Processed: $totalGroups</p>
        <p>Inactive Groups Found: $($inactiveGroups.Count)</p>
        <p>Email Activity Window: Last $MessageTraceDays days ($(Get-Date $startDate -Format 'yyyy-MM-dd') to $(Get-Date $endDate -Format 'yyyy-MM-dd'))</p>
    </div>
"@

    foreach ($group in $inactiveGroups) {
        $htmlContent += @"
    <div class="group-details">
        <h3>$($group.Name)</h3>
        <table>
            <tr>
                <th style="width: 200px;">Property</th>
                <th>Value</th>
            </tr>
            <tr><td>Email Address</td><td>$($group.EmailAddress)</td></tr>
            <tr><td>Member Count</td><td>$($group.MemberCount)</td></tr>
            <tr><td>Last Modified</td><td>$($group.LastModified)</td></tr>
            <tr><td>Created</td><td>$($group.Created)</td></tr>
            <tr><td>Last Email Received</td><td>$($group.LastEmailReceived)</td></tr>
            <tr><td>Last Email Sent</td><td>$($group.LastEmailSent)</td></tr>
            <tr><td>Hidden From GAL</td><td>$($group.HiddenFromGAL)</td></tr>
        </table>

        <h4>Members</h4>
        <div class="members-list">
            <pre>$($group.Members)</pre>
        </div>

        <h4>Owners</h4>
        <div class="members-list">
            <pre>$($group.Owners)</pre>
        </div>
    </div>
"@
    }

    $htmlContent += @"
    <div class="summary">
        <h3>Report Notes:</h3>
        <ul>
            <li>This report shows distribution groups with no activity in the last $InactiveDays days.</li>
            <li>Email activity data is limited to the last $MessageTraceDays days due to Exchange Online limitations.</li>
            <li>Groups with no recorded email activity may still be in use for other purposes.</li>
            <li>Please verify group usage before taking any action.</li>
            <li>Member and owner lists are current as of report generation time.</li>
        </ul>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $htmlReport -Encoding UTF8
    Write-Host "HTML report saved to: $htmlReport" -ForegroundColor Green

    # Display summary
    Write-Host "`nScan Complete!" -ForegroundColor Green
    Write-Host "Total Groups Processed: $totalGroups" -ForegroundColor Cyan
    Write-Host "Inactive Groups Found: $($inactiveGroups.Count)" -ForegroundColor Yellow
    Write-Host "CSV Report exported to: $csvPath" -ForegroundColor Green
    Write-Host "HTML Report exported to: $htmlReport" -ForegroundColor Green

    # Optional: Display preview of results
    Write-Host "`nPreview of Inactive Groups:" -ForegroundColor Cyan
    $inactiveGroups | Select-Object Name, EmailAddress, MemberCount, LastEmailReceived, LastEmailSent | Format-Table -AutoSize
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    throw
}
finally {
    # Disconnecting from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    
    Stop-Transcript
}
