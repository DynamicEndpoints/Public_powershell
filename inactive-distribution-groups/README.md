# Distribution Group Activity Analyzer

## Overview

The Distribution Group Activity Analyzer is a PowerShell tool designed to help Exchange Online administrators identify and manage inactive distribution groups. This tool scans your Exchange Online environment, identifies distribution groups with no recent activity, and generates detailed reports to assist with cleanup and management decisions.

## Features

- **Comprehensive Analysis**: Scans all distribution groups in your Exchange Online environment (or a filtered subset)
- **Activity Detection**: Identifies groups with no email activity within a specified timeframe
- **Detailed Reporting**: Generates both CSV and professional HTML reports with actionable insights
- **Group Details**: Collects extensive information about each inactive group including:
  - Membership details and count
  - Group owners
  - Email activity (last sent/received)
  - Creation and modification dates
  - Various group properties and settings
- **Executive Summary**: Provides key metrics and recommendations for management
- **Customizable Parameters**: Configurable inactivity thresholds and scanning options

## Requirements

- PowerShell 5.1 or later
- Exchange Online PowerShell V2 module (`Install-Module -Name ExchangeOnlineManagement`)
- Exchange Online account with appropriate permissions (Global Admin, Exchange Admin, or similar)
- Internet connectivity to connect to Exchange Online

## Installation

1. Download the `inactive_dl_groups_v3.ps1` script to your local machine
2. Ensure you have the Exchange Online PowerShell V2 module installed:
   ```powershell
   Install-Module -Name ExchangeOnlineManagement -Force
   ```

## Usage

### Basic Usage

Run the script with default parameters:

```powershell
.\inactive_dl_groups_v3.ps1
```

This will:
- Use a 90-day inactivity threshold
- Look back 10 days for message activity
- Save reports to your Documents folder

### Advanced Usage

```powershell
.\inactive_dl_groups_v3.ps1 -InactiveDays 180 -MessageTraceDays 30 -ReportPath "C:\Reports" -DomainFilter "contoso.com","fabrikam.com" -GroupFilter "Marketing*","Sales*"
```

### Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-InactiveDays` | Number of days to consider a group inactive | 90 |
| `-MessageTraceDays` | Number of days to look back for email activity | 10 |
| `-ReportPath` | Path to save the reports | Documents folder |
| `-GroupFilter` | Filter groups by name pattern (accepts wildcards) | All groups |
| `-DomainFilter` | Filter groups by email domain | All domains |

## Output

The script generates three types of output:

1. **Log File**: A transcript of the script execution (saved to the report path)
2. **CSV Report**: Detailed data about inactive groups in CSV format for further analysis
3. **HTML Report**: Professional, management-ready report with:
   - Executive summary and key metrics
   - Detailed information about each inactive group
   - Visual indicators for critical information
   - Recommendations for handling inactive groups

## Interpreting the Results

The HTML report is divided into several sections:

- **Executive Summary**: Shows total groups analyzed, number of inactive groups, and inactivity rate
- **Inactive Groups Detail**: Provides comprehensive information about each inactive group
- **Recommendations**: Suggests actions to take based on the findings

Groups with the following characteristics deserve special attention:
- **Hidden from GAL**: May indicate groups that were intentionally hidden before deprecation
- **No Updates in 1+ Year**: Groups that have not been modified in a long time
- **No members**: Empty groups that may be obsolete
- **Large member count with no activity**: Potentially obsolete groups that still have many members

## Notes and Limitations

- Message trace data is limited to 10 days by default (Exchange Online limitation)
- The script authenticates to Exchange Online and will prompt for credentials
- Running the script may take significant time in environments with many distribution groups
- Group activity is determined by:
  1. Message trace data (emails sent/received)
  2. Modification dates of the group
- Some inactive groups may still be serving valid purposes not related to email activity

## Best Practices

1. **Verify Before Action**: Always verify with group owners before deleting groups
2. **Document Decisions**: Document any decisions to keep seemingly inactive groups
3. **Regular Audits**: Run this script periodically (quarterly recommended) to maintain a clean environment
4. **Incremental Cleanup**: Address the most obvious inactive groups first (oldest, empty, etc.)
5. **Test Mode**: Consider using `-WhatIf` with any removal actions after identifying inactive groups

## Troubleshooting

**Q: The script is running slowly**  
A: Reduce the scope using `-GroupFilter` or `-DomainFilter` to analyze subsets of groups.

**Q: I get authentication errors**  
A: Ensure you have the correct permissions and the latest Exchange Online Management module.

**Q: Some groups show no activity but I know they're used**  
A: Some groups may be used for purposes other than email. Message trace data is also limited to recent history.

## Version History

- **v1.0.0** (March 2025) - Initial release
- **v1.1.0** (March 2025) - Added enhanced HTML reporting and executive summary
- **v1.2.0** (March 2025) - Added domain and group filtering options

## License

This script is provided as-is under the MIT License.

## Author

Created by Your IT Department  
For support, contact your IT Service Desk.# Distribution Group Activity Analyzer
