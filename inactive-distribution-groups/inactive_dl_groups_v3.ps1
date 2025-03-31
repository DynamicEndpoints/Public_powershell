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

    # Apply domain filtering if specified using hash table for better performance
    if ($DomainFilter) {
        # Create domain lookup hash table for O(1) lookups
        $domainLookup = @{}
        foreach ($domain in $DomainFilter) {
            $domainLookup[$domain] = $true
        }
        
        # Filter groups using hash table lookup instead of Where-Object
        $filteredGroups = @()
        foreach ($group in $groups) {
            $groupDomain = ($group.PrimarySmtpAddress -split '@')[1]
            if ($domainLookup.ContainsKey($groupDomain)) {
                $filteredGroups += $group
            }
        }
        $groups = $filteredGroups
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
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Distribution Groups Activity Report</title>
    <style>
        :root {
            --primary-color: #0078D4; /* Microsoft blue */
            --secondary-color: #2b579a;
            --accent-color: #f3f3f3;
            --success-color: #107c10; /* Microsoft green */
            --warning-color: #ff8c00; /* Warning orange */
            --danger-color: #d13438; /* Microsoft red */
            --text-color: #333;
            --border-color: #ddd;
        }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 0;
            color: var(--text-color);
            line-height: 1.6;
            background-color: #fafafa;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        .header {
            background: var(--primary-color);
            color: white;
            padding: 20px;
            border-radius: 8px 8px 0 0;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
        }
        .header p {
            margin: 5px 0 0 0;
            opacity: 0.9;
        }
        .report-info {
            display: flex;
            justify-content: space-between;
            margin-top: 15px;
            font-size: 14px;
        }
        .summary-box {
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .summary-box h2 {
            margin-top: 0;
            color: var(--primary-color);
            border-bottom: 2px solid var(--accent-color);
            padding-bottom: 10px;
        }
        .metric-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        .metric-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-top: 4px solid var(--primary-color);
            transition: transform 0.2s ease;
        }
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        .metric-value {
            font-size: 32px;
            font-weight: bold;
            color: var(--primary-color);
            margin: 10px 0;
        }
        .metric-label {
            font-size: 14px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .section-title {
            color: var(--secondary-color);
            border-bottom: 2px solid var(--accent-color);
            padding-bottom: 10px;
            margin: 30px 0 20px 0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        th {
            background: var(--secondary-color);
            color: white;
            padding: 12px 15px;
            text-align: left;
            font-weight: 600;
        }
        td {
            padding: 12px 15px;
            border-bottom: 1px solid var(--border-color);
        }
        tr:nth-child(even) {
            background: var(--accent-color);
        }
        tr:hover {
            background-color: rgba(0,120,212,0.05);
        }
        .group-container {
            margin-bottom: 30px;
        }
        .group-details {
            background: white;
            border-radius: 8px;
            padding: 0;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .group-header {
            background: var(--secondary-color);
            color: white;
            padding: 15px 20px;
        }
        .group-header h3 {
            margin: 0;
            font-size: 18px;
        }
        .group-header small {
            display: block;
            margin-top: 5px;
            opacity: 0.9;
        }
        .group-content {
            padding: 20px;
        }
        .group-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        .info-box {
            background: var(--accent-color);
            border-radius: 6px;
            padding: 15px;
        }
        .info-box h4 {
            margin-top: 0;
            color: var(--secondary-color);
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 8px;
            margin-bottom: 10px;
        }
        .members-list {
            max-height: 200px;
            overflow-y: auto;
            padding: 10px;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            background: white;
            font-size: 13px;
        }
        .alert {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            color: #856404;
            padding: 12px 15px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
            text-transform: uppercase;
            margin-left: 10px;
        }
        .badge-warning {
            background: var(--warning-color);
            color: white;
        }
        .badge-danger {
            background: var(--danger-color);
            color: white;
        }
        .footer {
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: #666;
            border-top: 1px solid var(--border-color);
            font-size: 12px;
        }
        .chart-container {
            height: 300px;
            margin: 20px 0;
        }
        @media print {
            body {
                background: white;
                font-size: 12pt;
            }
            .container {
                max-width: 100%;
                padding: 10px;
            }
            .metric-card:hover {
                transform: none;
                box-shadow: none;
            }
            .group-details {
                page-break-inside: avoid;
                break-inside: avoid;
            }
            .members-list {
                max-height: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Distribution Groups Activity Report</h1>
            <p>$((Get-OrganizationConfig).DisplayName)</p>
            <div class="report-info">
                <span>Generated: $(Get-Date -Format "MMMM d, yyyy h:mm tt")</span>
                <span>Report Period: $(Get-Date $startDate -Format 'MMM d, yyyy') - $(Get-Date $endDate -Format 'MMM d, yyyy')</span>
            </div>
        </div>

        <div class="summary-box">
            <h2>Executive Summary</h2>
            <div class="metric-container">
                <div class="metric-card">
                    <div class="metric-label">Total Groups</div>
                    <div class="metric-value">$totalGroups</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Inactive Groups</div>
                    <div class="metric-value">$($inactiveGroups.Count)</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Inactivity Rate</div>
                    <div class="metric-value">$([math]::Round(($inactiveGroups.Count / $totalGroups) * 100, 1))%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Analysis Window</div>
                    <div class="metric-value">$MessageTraceDays</div>
                    <div style="font-size: 12px;">days</div>
                </div>
            </div>

            <div class="alert">
                <strong>Note:</strong> This report identifies distribution groups with no activity in the last $InactiveDays days. 
                Email activity data is limited to the last $MessageTraceDays days due to Exchange Online limitations.
                Please verify before taking any action, as groups may be used for purposes other than email.
            </div>
        </div>

        <h2 class="section-title">Inactive Distribution Groups Detail</h2>
        <div class="group-container">
"@

    foreach ($group in $inactiveGroups) {
        # Calculate days since creation and last modification
        $daysSinceCreation = [math]::Round(((Get-Date) - $group.Created).TotalDays)
        $daysSinceModified = [math]::Round(((Get-Date) - $group.LastModified).TotalDays)
        
        # Determine status badges based on group characteristics
        $statusBadge = ""
        if ($group.HiddenFromGAL) {
            $statusBadge += "<span class='badge badge-warning'>Hidden from GAL</span>"
        }
        if ($daysSinceModified -gt 365) {
            $statusBadge += "<span class='badge badge-danger'>No Updates in 1+ Year</span>"
        }
        
        $htmlContent += @"
            <div class="group-details">
                <div class="group-header">
                    <h3>$($group.Name)$statusBadge</h3>
                    <small>$($group.EmailAddress)</small>
                </div>
                
                <div class="group-content">
                    <div class="group-grid">
                        <div class="info-box">
                            <h4>Group Information</h4>
                            <table>
                                <tr><td><strong>Created</strong></td><td>$($group.Created) <small>($daysSinceCreation days ago)</small></td></tr>
                                <tr><td><strong>Last Modified</strong></td><td>$($group.LastModified) <small>($daysSinceModified days ago)</small></td></tr>
                                <tr><td><strong>Member Count</strong></td><td>$($group.MemberCount)</td></tr>
                                <tr><td><strong>Hidden From GAL</strong></td><td>$($group.HiddenFromGAL)</td></tr>
                                <tr><td><strong>Authentication Required</strong></td><td>$($group.RequireSenderAuthenticationEnabled)</td></tr>
                            </table>
                        </div>
                        <div class="info-box">
                            <h4>Email Activity</h4>
                            <table>
                                <tr><td><strong>Last Email Received</strong></td><td>$($group.LastEmailReceived)</td></tr>
                                <tr><td><strong>Last Received From</strong></td><td>$($group.LastEmailReceivedFrom)</td></tr>
                                <tr><td><strong>Last Email Sent</strong></td><td>$($group.LastEmailSent)</td></tr>
                                <tr><td><strong>Last Sent To</strong></td><td>$($group.LastEmailSentTo)</td></tr>
                                <tr><td><strong>Activity Status</strong></td><td>No activity for at least $InactiveDays days</td></tr>
                            </table>
                        </div>
                    </div>

                    <div class="group-grid">
                        <div class="info-box">
                            <h4>Members ($($group.MemberCount))</h4>
                            <div class="members-list">
                                <pre style="margin: 0;">$($group.Members)</pre>
                            </div>
                        </div>
                        <div class="info-box">
                            <h4>Owners</h4>
                            <div class="members-list">
                                <pre style="margin: 0;">$($group.Owners)</pre>
                            </div>
                        </div>
                    </div>
                    
                    <div class="info-box">
                        <h4>Additional Properties</h4>
                        <table>
                            <tr><td><strong>Additional Email Addresses</strong></td><td>$($group.AllEmailAddresses)</td></tr>
                            <tr><td><strong>Accept Messages Only From</strong></td><td>$($group.AcceptMessagesOnlyFrom)</td></tr>
                            <tr><td><strong>Accept Messages Only From DL Members</strong></td><td>$($group.AcceptMessagesOnlyFromDLMembers)</td></tr>
                            <tr><td><strong>Custom Attribute 1</strong></td><td>$($group.CustomAttribute1)</td></tr>
                            <tr><td><strong>Custom Attribute 2</strong></td><td>$($group.CustomAttribute2)</td></tr>
                            <tr><td><strong>Notes</strong></td><td>$($group.Notes)</td></tr>
                        </table>
                    </div>
                </div>
            </div>
"@
    }

    $htmlContent += @"
        </div>

        <div class="summary-box">
            <h2>Recommendations</h2>
            <p>Based on the analysis of $totalGroups distribution groups, we found $($inactiveGroups.Count) groups with no apparent activity in the last $InactiveDays days. Consider these actions:</p>
            
            <ol>
                <li><strong>Verify Usage:</strong> Confirm with owners if these groups are still needed</li>
                <li><strong>Archive or Remove:</strong> Decommission unused groups to improve directory hygiene</li>
                <li><strong>Document:</strong> For groups that must be retained, document the business justification</li>
                <li><strong>Review Membership:</strong> For groups being kept, verify that membership is current</li>
            </ol>
            
            <p>Regular maintenance of distribution groups helps maintain security, reduces clutter, and improves overall management of your Exchange environment.</p>
        </div>
        
        <div class="footer">
            <p>Report generated by Exchange Online Distribution Group Activity Analyzer | &copy; $(Get-Date -Format "yyyy") IT Department</p>
            <p>Contact the IT Service Desk for questions or assistance with managing distribution groups</p>
        </div>
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
