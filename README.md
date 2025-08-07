# PRTG User Report

PowerShell script for retrieving PRTG user information with multi-server support and comprehensive Excel reporting.

## Script

### Get-PRTGUserLogins.ps1
Advanced script that scrapes PRTG server web interfaces to create comprehensive user reports. Supports multiple servers with interactive prompts, creating Excel reports with separate worksheets for each server.

**Usage:**
```powershell
# Interactive mode (recommended)
./Get-PRTGUserLogins.ps1

# Custom output file
./Get-PRTGUserLogins.ps1 -OutputPath "MyCompanyPRTG.xlsx"

# Specify first server details
./Get-PRTGUserLogins.ps1 -Server "https://prtg1.company.com" -Username "admin" -Password "mypassword"
```


## Features

### Interactive Mode
- Processes first server using provided parameters
- After completion, prompts: "Add another PRTG server to this report? (Y/N)"
- For each additional server, collects:
  - Server URL
  - Username  
  - Password
- Continues until user chooses "No"

### Excel Output (.xlsx)
- **Excel output** with multiple worksheets (one per server)
- **Each server gets its own worksheet** named after the server hostname
- **Auto-formatted headers** (bold, frozen top row)
- **Auto-filter enabled** for easy sorting/filtering
- **Professional appearance** with optimal column widths
- **Automatic Python/openpyxl installation** for cross-platform compatibility


### Data Columns
- **User Name**: User's display name in PRTG
- **Account Status**: Active or Paused status (correctly extracted)
- **Last Login Date**: Last login date in M/d/yyyy format (e.g., "8/5/2025"), or "(has not logged in yet)" for users who have never logged in

### Summary Report
- **Total users** across all servers
- **Per-server breakdown** with user counts and status
- **Processing results** for each server

## Permissions

- **User List Access**: Most PRTG users can access the general user list
- **Detailed Login Info**: Requires administrative permissions to view individual user login details
- **API Access**: User account must have API access enabled in PRTG settings

## Technical Notes

- Uses **web scraping** for reliable user enumeration (more consistent than JSON API for user data)
- Handles **special characters in passwords** with proper URL encoding
- **SSL certificate validation** is disabled to support self-signed certificates
- **Cross-platform compatible** with Windows PowerShell and PowerShell Core

## Requirements

- **PowerShell 5.1+** or **PowerShell Core 6+**
- **Python 3.x** (automatically installed openpyxl for Excel export)
- **PRTG API access** enabled for the user account
- **Administrative permissions** in PRTG to view user login details