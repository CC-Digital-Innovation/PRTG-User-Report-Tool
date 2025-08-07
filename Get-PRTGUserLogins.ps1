#!/usr/bin/env pwsh
#requires -version 4.0

<#
.SYNOPSIS
    Gets PRTG user information from multiple servers by scraping their web interfaces
    
.DESCRIPTION
    This script scrapes PRTG's web interface to get user information from one or more servers.
    Uses web scraping approach since it's more reliable than the JSON API for user data.
    Fixes the authentication issues with special characters in passwords.
    Supports multi-server mode with interactive prompts or single server mode.
    
.PARAMETER Server
    PRTG server URL for the first server (e.g., https://prtg.example.com)
    
.PARAMETER Username
    PRTG username for authentication on the first server
    
.PARAMETER Password
    PRTG password for authentication on the first server
    
.PARAMETER OutputPath
    Path for output file (defaults to PRTG_MultiServer_<timestamp>.xlsx)
    
.PARAMETER ExportToExcel
    Force Excel export (automatically enabled if OutputPath ends with .xlsx)
    
.EXAMPLE
    ./Get-PRTGUserLogins.ps1 -Server "https://prtg.example.com" -Username "admin" -Password "mypassword"
    
.EXAMPLE
    .EXAMPLE  
    ./Get-PRTGUserLogins.ps1 -OutputPath "MyReport.xlsx"
    
.EXAMPLE
    # Interactive mode - script will prompt for server details
    ./Get-PRTGUserLogins.ps1
#>

param(
    [string]$Server,
    [string]$Username, 
    [string]$Password,
    [string]$OutputPath = "PRTG_MultiServer_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    
    [switch]$ExportToExcel
)

Write-Host "`n=== PRTG User Report (Web Scraping Method) ===" -ForegroundColor Cyan
Write-Host "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Cyan

# Python-based Excel export functions for macOS compatibility
function Export-ExcelPython {
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,
        
        [Parameter(Mandatory)]
        [string]$Path,
        
        [string]$WorksheetName = "Sheet1",
        
        [switch]$FreezeTopRow,
        
        [switch]$BoldTopRow,
        
        [switch]$AutoFilter,
        
        [hashtable]$ColumnWidths = @{}
    )
    
    Write-Host "Exporting to Excel using Python..." -ForegroundColor Yellow
    
    # Check if Python is available - try different commands based on platform
    $pythonCmd = $null
    $pythonVersion = $null
    
    # Try different Python commands in order of preference
    $pythonCommands = @('python3', 'python', 'py -3')
    
    foreach ($cmd in $pythonCommands) {
        try {
            $pythonVersion = if ($cmd -eq 'py -3') { 
                Invoke-Expression "$cmd --version" 2>&1 
            } else { 
                & $cmd --version 2>&1 
            }
            
            if ($pythonVersion -match "Python 3\.\d+") {
                $pythonCmd = $cmd
                Write-Host "Using $pythonVersion with command: $pythonCmd" -ForegroundColor Green
                break
            }
        } catch {
            # Continue to next command
        }
    }
    
    if (-not $pythonCmd) {
        $installMsg = if ($IsWindows -or $env:OS -eq "Windows_NT") {
            "Install from: https://www.python.org/downloads/ or Microsoft Store"
        } else {
            "Install with: brew install python3"
        }
        throw "Python 3 is required but not found. $installMsg"
    }
    
    # Check if openpyxl is installed
    $checkOpenpyxl = try {
        if ($pythonCmd -eq 'py -3') {
            Invoke-Expression "$pythonCmd -c 'import openpyxl; print(`"openpyxl available`")'" -ErrorAction SilentlyContinue
        } else {
            & $pythonCmd -c "import openpyxl; print('openpyxl available')" 2>&1
        }
    } catch {
        "openpyxl not found"
    }
    
    if ($checkOpenpyxl -notlike "*openpyxl available*") {
        Write-Host "Installing openpyxl..." -ForegroundColor Yellow
        if ($pythonCmd -eq 'py -3') {
            Invoke-Expression "$pythonCmd -m pip install --user openpyxl"
        } else {
            & $pythonCmd -m pip install --user --break-system-packages openpyxl
        }
    }
    
    # Convert PowerShell objects to JSON for Python
    $jsonData = $Data | ConvertTo-Json -Depth 10
    $tempJson = [System.IO.Path]::GetTempFileName() + ".json"
    $jsonData | Out-File -FilePath $tempJson -Encoding UTF8
    
    # Convert column widths to Python dict format (convert hashtable to custom object first)
    $columnWidthsJson = if ($ColumnWidths.Count -gt 0) {
        $widthsObject = [PSCustomObject]@{}
        foreach ($key in $ColumnWidths.Keys) {
            $widthsObject | Add-Member -MemberType NoteProperty -Name $key.ToString() -Value $ColumnWidths[$key]
        }
        $widthsObject | ConvertTo-Json -Compress
    } else {
        "{}"
    }
    
    # Create Python script
    $pythonScript = @"
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.filters import AutoFilter
import sys
import os

# Read JSON data
with open('$tempJson', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Parse column widths
column_widths = json.loads('$columnWidthsJson')

# Create workbook and worksheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '$WorksheetName'

if not data:
    print("No data to export")
    sys.exit(1)

# Get headers from first object
headers = list(data[0].keys())

# Add headers
for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    
    # Bold headers if requested
    if '$BoldTopRow' == 'True':
        cell.font = Font(bold=True)

# Add data rows
for row_idx, item in enumerate(data, 2):
    for col_idx, header in enumerate(headers, 1):
        value = item.get(header, '')
        cell = ws.cell(row=row_idx, column=col_idx)
        
        # Handle date formatting for Last Login Date column
        if header == 'Last Login Date' and value and value not in ['Not found', 'Error', '(has not logged in yet)']:
            try:
                from datetime import datetime
                # Try to parse the date string and convert to Excel date
                if '/' in str(value):
                    # Parse M/d/yyyy format
                    date_obj = datetime.strptime(str(value), '%m/%d/%Y')
                    cell.value = date_obj
                    # Apply date format to the cell
                    cell.number_format = 'M/D/YYYY'
                else:
                    cell.value = value
            except:
                # If date parsing fails, use original value
                cell.value = value
        else:
            cell.value = value

# Set column widths
for col_idx, width in column_widths.items():
    col_letter = openpyxl.utils.get_column_letter(int(col_idx))
    ws.column_dimensions[col_letter].width = width

# Freeze top row if requested
if '$FreezeTopRow' == 'True':
    ws.freeze_panes = 'A2'

# Add AutoFilter if requested
if '$AutoFilter' == 'True':
    ws.auto_filter = AutoFilter(ref=f"A1:{openpyxl.utils.get_column_letter(len(headers))}{len(data)+1}")

# Save workbook
wb.save('$Path')
print(f"✓ Excel file created: $Path")
"@
    
    # Execute Python script
    try {
        $result = if ($pythonCmd -eq 'py -3') {
            Invoke-Expression "$pythonCmd -c `"$pythonScript`"" 2>&1
        } else {
            & $pythonCmd -c $pythonScript 2>&1
        }
        Write-Host $result -ForegroundColor Green
    } catch {
        throw "Failed to create Excel file: $_"
    } finally {
        # Clean up temp file
        if (Test-Path $tempJson) {
            Remove-Item $tempJson
        }
    }
}

function Add-ExcelWorksheetPython {
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,
        
        [Parameter(Mandatory)]
        [string]$Path,
        
        [Parameter(Mandatory)]
        [string]$WorksheetName,
        
        [switch]$FreezeTopRow,
        
        [switch]$BoldTopRow,
        
        [switch]$AutoFilter,
        
        [hashtable]$ColumnWidths = @{}
    )
    
    Write-Host "Adding worksheet '$WorksheetName' to Excel file..." -ForegroundColor Yellow
    
    # Convert PowerShell objects to JSON for Python
    $jsonData = $Data | ConvertTo-Json -Depth 10
    $tempJson = [System.IO.Path]::GetTempFileName() + ".json"
    $jsonData | Out-File -FilePath $tempJson -Encoding UTF8
    
    # Convert column widths to Python dict format (convert hashtable to custom object first)
    $columnWidthsJson = if ($ColumnWidths.Count -gt 0) {
        $widthsObject = [PSCustomObject]@{}
        foreach ($key in $ColumnWidths.Keys) {
            $widthsObject | Add-Member -MemberType NoteProperty -Name $key.ToString() -Value $ColumnWidths[$key]
        }
        $widthsObject | ConvertTo-Json -Compress
    } else {
        "{}"
    }
    
    # Create Python script to add worksheet
    $pythonScript = @"
import json
import openpyxl
from openpyxl.styles import Font
from openpyxl.worksheet.filters import AutoFilter
import sys
import os

# Check if Excel file exists
if not os.path.exists('$Path'):
    print("Excel file does not exist, creating new one...")
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
else:
    # Load existing workbook
    wb = openpyxl.load_workbook('$Path')

# Read JSON data
with open('$tempJson', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Parse column widths
column_widths = json.loads('$columnWidthsJson')

# Create new worksheet
ws = wb.create_sheet('$WorksheetName')

if not data:
    print("No data to export")
    sys.exit(1)

# Get headers from first object
headers = list(data[0].keys())

# Add headers
for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    
    # Bold headers if requested
    if '$BoldTopRow' == 'True':
        cell.font = Font(bold=True)

# Add data rows
for row_idx, item in enumerate(data, 2):
    for col_idx, header in enumerate(headers, 1):
        value = item.get(header, '')
        cell = ws.cell(row=row_idx, column=col_idx)
        
        # Handle date formatting for Last Login Date column
        if header == 'Last Login Date' and value and value not in ['Not found', 'Error', '(has not logged in yet)']:
            try:
                from datetime import datetime
                # Try to parse the date string and convert to Excel date
                if '/' in str(value):
                    # Parse M/d/yyyy format
                    date_obj = datetime.strptime(str(value), '%m/%d/%Y')
                    cell.value = date_obj
                    # Apply date format to the cell
                    cell.number_format = 'M/D/YYYY'
                else:
                    cell.value = value
            except:
                # If date parsing fails, use original value
                cell.value = value
        else:
            cell.value = value

# Set column widths
for col_idx, width in column_widths.items():
    col_letter = openpyxl.utils.get_column_letter(int(col_idx))
    ws.column_dimensions[col_letter].width = width

# Freeze top row if requested
if '$FreezeTopRow' == 'True':
    ws.freeze_panes = 'A2'

# Add AutoFilter if requested
if '$AutoFilter' == 'True':
    ws.auto_filter = AutoFilter(ref=f"A1:{openpyxl.utils.get_column_letter(len(headers))}{len(data)+1}")

# Save workbook
wb.save('$Path')
print(f"✓ Worksheet '$WorksheetName' added to Excel file")
"@
    
    # Execute Python script
    try {
        # Check if Python is available - try different commands based on platform
        $pythonCmd = $null
        $pythonCommands = @('python3', 'python', 'py -3')
        
        foreach ($cmd in $pythonCommands) {
            try {
                $pythonVersion = if ($cmd -eq 'py -3') { 
                    Invoke-Expression "$cmd --version" 2>&1 
                } else { 
                    & $cmd --version 2>&1 
                }
                
                if ($pythonVersion -match "Python 3\.\d+") {
                    $pythonCmd = $cmd
                    break
                }
            } catch {
                # Continue to next command
            }
        }
        
        if (-not $pythonCmd) {
            throw "Python 3 is required but not found"
        }
        
        $result = if ($pythonCmd -eq 'py -3') {
            Invoke-Expression "$pythonCmd -c `"$pythonScript`"" 2>&1
        } else {
            & $pythonCmd -c $pythonScript 2>&1
        }
        Write-Host $result -ForegroundColor Green
    } catch {
        throw "Failed to add worksheet to Excel file: $_"
    } finally {
        # Clean up temp file
        if (Test-Path $tempJson) {
            Remove-Item $tempJson
        }
    }
}

Write-Host "Using Python-based Excel export for macOS compatibility..." -ForegroundColor Yellow

# PowerShell 6+ SSL handling
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck'] = $true
    $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck'] = $true
}

# Function to get PRTG users from a single server
function Get-PRTGServerUsers {
    param(
        [string]$ServerUrl,
        [string]$Username,
        [string]$Password
    )
    
    $result = @{
        Users = @()
        ServerName = ""
        Success = $false
        ErrorMessage = ""
    }
    
    try {
        # Clean up server URL
        $ServerUrl = $ServerUrl.TrimEnd('/')
        if ($ServerUrl -notmatch '^https?://') {
            $ServerUrl = "https://$ServerUrl"
            Write-Host "Note: Added https:// prefix to server URL" -ForegroundColor Yellow
        }
        
        $result.ServerName = $ServerUrl -replace '^https?://', ''
        
        Write-Host "`nProcessing server: $ServerUrl" -ForegroundColor Cyan
        Write-Host "Username: $Username" -ForegroundColor Gray
        
        # Get passhash for authentication - with proper URL encoding for special characters
        Write-Host "`nStep 1: Getting authentication token..." -ForegroundColor Yellow
        
        # Use [System.Uri] for better URL encoding that handles special characters correctly
        $encodedUser = [System.Uri]::EscapeDataString($Username)
        $encodedPass = [System.Uri]::EscapeDataString($Password)
        
        $passhashUrl = "$ServerUrl/api/getpasshash.htm?username=$encodedUser&password=$encodedPass"
        $passhash = (Invoke-RestMethod -Uri $passhashUrl -Method Get).ToString()
        Write-Host "✓ Authentication successful" -ForegroundColor Green
        
        # Get user list from web interface
        Write-Host "`nStep 2: Retrieving user list..." -ForegroundColor Yellow
        $userlistUrl = "$ServerUrl/controls/userlist.htm?count=9000&username=$Username&passhash=$passhash"
        $userlistPage = Invoke-WebRequest -Uri $userlistUrl -UseBasicParsing
        
        Write-Host "✓ Retrieved user list page ($(($userlistPage.Content).Length) bytes)" -ForegroundColor Green
        
        # Extract user IDs and names from HTML
        Write-Host "`nStep 3: Parsing user data..." -ForegroundColor Yellow
        
        $idMatches = [regex]::Matches($userlistPage.Content, 'edituser\.htm\?id=(\d+)')
        $nameMatches = [regex]::Matches($userlistPage.Content, 'edituser\.htm\?id=\d+">([^<]+)<')
        
        Write-Host "Found $($idMatches.Count) user IDs and $($nameMatches.Count) user names" -ForegroundColor Gray
        
        if ($idMatches.Count -eq 0) {
            $result.ErrorMessage = "No users found. This could indicate permission issues or unexpected HTML structure."
            return $result
        }
        
        # Get detailed login information for each user
        Write-Host "`nStep 4: Retrieving detailed login information..." -ForegroundColor Yellow
        
        $userlogin_regex = '<div class="readonlyproperty" >(.+?)</div>'
        $userstatus_regex = '<input[^>]*name="active[^"]*"[^>]*value="([^"]*)"'
        $usergroup_regex = '<select[^>]*name="[^"]*primarygroup[^"]*"[^>]*>.*?<option[^>]*selected[^>]*>([^<]+)</option>'
        
        [string[]]$userlogin_times = $null
        [string[]]$user_statuses = $null
        [string[]]$user_groups = $null
        
        $progressCount = 0
        foreach ($match in $idMatches) {
            $userId = $match.Groups[1].Value
            $progressCount++
            Write-Progress -Activity "Getting user details" -Status "Processing user $progressCount of $($idMatches.Count)" -PercentComplete (($progressCount / $idMatches.Count) * 100)
            
            $userlogin_url = "$ServerUrl/controls/edituser.htm?id=$userId&username=$Username&passhash=$passhash"
            
            try {
                $userlogin = Invoke-WebRequest -Uri $userlogin_url -UseBasicParsing
                
                # Extract last login time and format as date only
                $userlogin_times_raw = [regex]::Matches($userlogin.Content, $userlogin_regex)
                if ($userlogin_times_raw.Count -gt 0) {
                    $rawLogin = $userlogin_times_raw[0].Groups[1].Value
                    
                    # Convert to date-only format if it's a valid timestamp
                    if ($rawLogin -match '\d+/\d+/\d+ \d+:\d+:\d+ [AP]M') {
                        try {
                            $loginDate = [DateTime]::Parse($rawLogin)
                            $userlogin_times += $loginDate.ToString('M/d/yyyy')
                        } catch {
                            $userlogin_times += $rawLogin  # Keep original if parsing fails
                        }
                    } elseif ($rawLogin -match '\d+/\d+/\d+') {
                        try {
                            $loginDate = [DateTime]::Parse($rawLogin)
                            $userlogin_times += $loginDate.ToString('M/d/yyyy')
                        } catch {
                            $userlogin_times += $rawLogin  # Keep original if parsing fails
                        }
                    } else {
                        $userlogin_times += $rawLogin  # Keep as-is for "(has not logged in yet)" etc.
                    }
                } else {
                    $userlogin_times += "Not found"
                }
                
                # Extract user status (active/paused) - look for the CHECKED radio button
                $checkedStatusRegex = '<input[^>]*name="active[^"]*"[^>]*value="([^"]*)"[^>]*checked'
                $checkedStatusMatch = [regex]::Matches($userlogin.Content, $checkedStatusRegex)
                
                
                if ($checkedStatusMatch.Count -gt 0) {
                    $statusValue = $checkedStatusMatch[0].Groups[1].Value
                    
                    # Convert PRTG status values to readable format
                    if ($statusValue -match '^(-1|1|true)$') {
                        $user_statuses += "Active"
                    } elseif ($statusValue -match '^(0|false)$') {
                        $user_statuses += "Paused"
                    } else {
                        $user_statuses += $statusValue  # Keep original if unknown format
                    }
                } else {
                    # Fallback: try to find any active field and check its value
                    $fallbackMatch = [regex]::Matches($userlogin.Content, $userstatus_regex)
                    if ($fallbackMatch.Count -gt 0) {
                        $statusValue = $fallbackMatch[0].Groups[1].Value
                        if ($statusValue -match '^(-1|1|true)$') {
                            $user_statuses += "Active"
                        } elseif ($statusValue -match '^(0|false)$') {
                            $user_statuses += "Paused"
                        } else {
                            $user_statuses += $statusValue
                        }
                    } else {
                        $user_statuses += "Unknown"
                    }
                }
                
                # Extract primary group
                $groupMatch = [regex]::Match($userlogin.Content, $usergroup_regex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                if ($groupMatch.Success) {
                    $groupName = $groupMatch.Groups[1].Value.Trim()
                    $user_groups += $groupName
                } else {
                    # Try alternative pattern for readonly group display
                    $readonlyGroupRegex = '<td[^>]*>Primary Group</td>.*?<td[^>]*>([^<]+)</td>'
                    $readonlyMatch = [regex]::Match($userlogin.Content, $readonlyGroupRegex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                    if ($readonlyMatch.Success) {
                        $user_groups += $readonlyMatch.Groups[1].Value.Trim()
                    } else {
                        $user_groups += "Unknown"
                    }
                }
                
            } catch {
                Write-Warning "Failed to get details for user ID $userId : $_"
                $userlogin_times += "Error"
                $user_statuses += "Error"
                $user_groups += "Error"
            }
        }
        
        Write-Progress -Activity "Getting user details" -Completed
        
        # Create user objects with essential information including status and group
        if ($idMatches.Count -eq $nameMatches.Count -and $idMatches.Count -eq $userlogin_times.Count -and $idMatches.Count -eq $user_statuses.Count -and $idMatches.Count -eq $user_groups.Count) {
            for ($i = 0; $i -lt $idMatches.Count; $i++) {
                $result.Users += [PSCustomObject]@{
                    "User Name" = $nameMatches[$i].Groups[1].Value
                    "Primary Group" = $user_groups[$i]
                    "Account Status" = $user_statuses[$i]
                    "Last Login Date" = $userlogin_times[$i]
                }
            }
        } else {
            $result.ErrorMessage = "Data mismatch: Names=$($nameMatches.Count), LoginTimes=$($userlogin_times.Count), Statuses=$($user_statuses.Count), Groups=$($user_groups.Count)"
            return $result
        }
        
        Write-Host "✓ Processed $($result.Users.Count) users from $($result.ServerName)" -ForegroundColor Green
        $result.Success = $true
        
    } catch {
        $result.ErrorMessage = "Server processing failed: $_"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $result.ErrorMessage += " (HTTP $statusCode)"
            
            if ($statusCode -eq 401) {
                $result.ErrorMessage += " - Authentication failed. Check username/password and API access."
            }
        }
    }
    
    return $result
}

# Function to get user input for additional server
function Get-AdditionalServerInfo {
    Write-Host "`n" -NoNewline
    do {
        $addAnother = Read-Host "Add another PRTG server to this report? (Y/N)"
        $addAnother = $addAnother.Trim().ToUpper()
    } while ($addAnother -ne 'Y' -and $addAnother -ne 'N')
    
    if ($addAnother -eq 'N') {
        return $null
    }
    
    Write-Host "`n=== Additional PRTG Server ===" -ForegroundColor Cyan
    
    $serverUrl = Read-Host "Enter PRTG server URL (e.g., https://prtg2.example.com)"
    if ([string]::IsNullOrWhiteSpace($serverUrl)) {
        Write-Host "Invalid server URL. Skipping..." -ForegroundColor Red
        return $null
    }
    
    $username = Read-Host "Enter username"
    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host "Invalid username. Skipping..." -ForegroundColor Red
        return $null
    }
    
    # Try SecureString with proper BSTR handling for special characters
    # Using PtrToStringBSTR instead of PtrToStringAuto for cross-platform compatibility
    # and proper handling of special characters like backticks, quotes, dollar signs
    try {
        Write-Host "Enter password (input will be hidden): " -NoNewline -ForegroundColor Yellow
        $securePassword = Read-Host -AsSecureString
        
        # Use proper BSTR conversion with memory cleanup for cross-platform compatibility
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        try {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
        }
        finally {
            # Critical: Always free memory to prevent leaks
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
        
        if ([string]::IsNullOrWhiteSpace($plainPassword)) {
            Write-Host "Invalid password. Skipping..." -ForegroundColor Red
            return $null
        }
    } catch {
        # Fallback to plain text if SecureString fails
        Write-Host "`nSecureString conversion failed, falling back to plain text input" -ForegroundColor Yellow
        $plainPassword = Read-Host "Enter password (will be visible)"
        if ([string]::IsNullOrWhiteSpace($plainPassword)) {
            Write-Host "Invalid password. Skipping..." -ForegroundColor Red
            return $null
        }
    }
    
    return @{
        Server = $serverUrl
        Username = $username
        Password = $plainPassword
    }
}

# Function to get first server credentials
function Get-FirstServerInfo {
    Write-Host "=== First PRTG Server ===" -ForegroundColor Cyan
    
    if ([string]::IsNullOrWhiteSpace($Server)) {
        $serverUrl = Read-Host "Enter PRTG server URL (e.g., https://prtg.example.com)"
        if ([string]::IsNullOrWhiteSpace($serverUrl)) {
            throw "Server URL is required"
        }
    } else {
        $serverUrl = $Server
        Write-Host "Server: $serverUrl" -ForegroundColor Gray
    }
    
    if ([string]::IsNullOrWhiteSpace($Username)) {
        $username = Read-Host "Enter username"
        if ([string]::IsNullOrWhiteSpace($username)) {
            throw "Username is required"
        }
    } else {
        $username = $Username
        Write-Host "Username: $username" -ForegroundColor Gray
    }
    
    if ([string]::IsNullOrWhiteSpace($Password)) {
        # Try SecureString with proper BSTR handling for special characters
        # Using PtrToStringBSTR instead of PtrToStringAuto for cross-platform compatibility
        # and proper handling of special characters like backticks, quotes, dollar signs
        try {
            Write-Host "Enter password (input will be hidden): " -NoNewline -ForegroundColor Yellow
            $securePassword = Read-Host -AsSecureString
            
            # Use proper BSTR conversion with memory cleanup for cross-platform compatibility
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
            try {
                $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
            }
            finally {
                # Critical: Always free memory to prevent leaks
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            }
            
            if ([string]::IsNullOrWhiteSpace($plainPassword)) {
                throw "Password is required"
            }
        } catch {
            # Fallback to plain text if SecureString fails
            Write-Host "`nSecureString conversion failed, falling back to plain text input" -ForegroundColor Yellow
            $plainPassword = Read-Host "Enter password (will be visible)"
            if ([string]::IsNullOrWhiteSpace($plainPassword)) {
                throw "Password is required"
            }
        }
    } else {
        $plainPassword = $Password
        Write-Host "Password: [provided]" -ForegroundColor Gray
    }
    
    return @{
        Server = $serverUrl
        Username = $username
        Password = $plainPassword
    }
}

# Main processing logic
try {
    # Initialize collections for all server data
    $allServerData = @()
    $totalUsers = 0
    $excel = $null
    
    # Get first server credentials
    $firstServer = Get-FirstServerInfo
    
    # Process first server
    Write-Host "`n=== Processing Server 1 ===" -ForegroundColor Magenta
    $serverResult = Get-PRTGServerUsers -ServerUrl $firstServer.Server -Username $firstServer.Username -Password $firstServer.Password
    
    if ($serverResult.Success) {
        $allServerData += $serverResult
        $totalUsers += $serverResult.Users.Count
        
        # Export first server data to Excel
        Write-Host "Creating Excel workbook with first server data..." -ForegroundColor Yellow
        # Sanitize worksheet name - Excel has restrictions on worksheet names
        $worksheetName = $serverResult.ServerName -replace '[\\\/\?\*\[\]:]', '_'
        if ($worksheetName.Length -gt 31) {
            $worksheetName = $worksheetName.Substring(0, 31)
        }
        
        # Use Python-based Excel export with proper column widths
        Export-ExcelPython -Data $serverResult.Users -Path $OutputPath -WorksheetName $worksheetName -FreezeTopRow -BoldTopRow -AutoFilter -ColumnWidths @{1=35; 2=25; 3=15; 4=18}
    } else {
        Write-Error "Failed to process first server: $($serverResult.ErrorMessage)"
        Write-Host "Continuing with additional servers..." -ForegroundColor Yellow
    }
    
    # Process additional servers
    {
        $serverCount = 1
        
        while ($true) {
            $additionalServer = Get-AdditionalServerInfo
            if ($null -eq $additionalServer) {
                break  # User chose not to add more servers
            }
            
            $serverCount++
            Write-Host "`n=== Processing Server $serverCount ===" -ForegroundColor Magenta
            
            $serverResult = Get-PRTGServerUsers -ServerUrl $additionalServer.Server -Username $additionalServer.Username -Password $additionalServer.Password
            
            if ($serverResult.Success) {
                $allServerData += $serverResult
                $totalUsers += $serverResult.Users.Count
                
                # Add to Excel workbook as new worksheet
                Write-Host "Adding server data to Excel workbook..." -ForegroundColor Yellow
                # Sanitize worksheet name - Excel has restrictions on worksheet names
                $worksheetName = $serverResult.ServerName -replace '[\\\/\?\*\[\]:]', '_'
                if ($worksheetName.Length -gt 31) {
                    $worksheetName = $worksheetName.Substring(0, 31)
                }
                        
                # Add new worksheet to existing Excel file using Python
                Add-ExcelWorksheetPython -Data $serverResult.Users -Path $OutputPath -WorksheetName $worksheetName -FreezeTopRow -BoldTopRow -AutoFilter -ColumnWidths @{1=35; 2=25; 3=15; 4=18}
            } else {
                Write-Error "Failed to process server $($additionalServer.Server): $($serverResult.ErrorMessage)"
                Write-Host "Continuing with next server..." -ForegroundColor Yellow
            }
        }
    }
    
    # Check if Excel file was created
    if (Test-Path $OutputPath) {
        $fileInfo = Get-Item $OutputPath
        Write-Host "`n✓ Multi-server report exported to Excel: $OutputPath" -ForegroundColor Green
    } else {
        Write-Warning "Excel file was not created"
    }
    
    # Display final summary
    Write-Host "`n=== Final Summary ===" -ForegroundColor Cyan
    Write-Host "Servers Processed: $($allServerData.Count)" -ForegroundColor White
    Write-Host "Total Users Found: $totalUsers" -ForegroundColor White
    
    if ($allServerData.Count -gt 0) {
        Write-Host "`n=== Server Breakdown ===" -ForegroundColor Cyan
        foreach ($serverData in $allServerData) {
            $activeCount = ($serverData.Users | Where-Object { $_."Account Status" -eq "Active" }).Count
            $pausedCount = ($serverData.Users | Where-Object { $_."Account Status" -eq "Paused" }).Count
            Write-Host "  $($serverData.ServerName): $($serverData.Users.Count) users ($activeCount Active, $pausedCount Paused)" -ForegroundColor White
        }
    }
    
} catch {
    Write-Error "Script failed: $_"
}

Write-Host "`nEnd Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=== PRTG User Report Complete ===`n" -ForegroundColor Cyan