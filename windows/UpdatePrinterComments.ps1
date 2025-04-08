# This script updates the comments for a list of printers based on their names, IP addresses, and models.
# It reads the printer names from a specified file, retrieves the current printer information,
# constructs new comments in the format "PlanningUnit - Model - IP Address", and logs the changes.
# The script uses the Get-Printer and Set-Printer cmdlets to interact with the printers,
# and handles resolving hostnames to IP addresses.
#
# Usage:
# 1. Ensure the printer list file is located at the specified path in $PrinterListFile. Currently, it's set to "E:\Scripts\PrinterComments\printerlist.txt".
#    The file should contain one printer name per line.
# 2. Run the script to update the printer comments and log the changes.
# 3. The script will display the current and new comments for each printer and log the changes to a CSV file.
# 4. The script uses the -WhatIf parameter with Set-Printer to simulate the changes without applying them. If you want to apply the changes, remove the -WhatIf parameter.
#
#Created by: Mike Croucher
#Date: 2023-10-03
#Version: 1.0

# Check if the script is running with administrative privileges
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script requires administrative privileges. Please run it as an administrator."
    exit
}
# Set the execution policy to allow script execution (if needed)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

# Import the required module for printer management
Import-Module PrintManagement -ErrorAction Stop

# Check if the module was imported successfully
if (-Not (Get-Module -Name PrintManagement)) {
    Write-Host "Error: PrintManagement module could not be loaded. Ensure it is installed."
    exit
}
# Set the error action preference to stop on errors for better error handling
$ErrorActionPreference = "Stop"

# Define the directory for script files
$ScriptDir = "E:\Scripts\PrinterComments"

# Define file paths
$PrinterListFile = "$ScriptDir\printerlist.txt"
$LogFile = "$ScriptDir\Printer_Comment_Updates.csv"

# Verify the printer list file exists
if (-Not (Test-Path $PrinterListFile)) {
    Write-Host "Error: Printer list file not found at $PrinterListFile"
    exit
}

# Read the printer names from the file (one per line)
$PrinterNames = Get-Content $PrinterListFile | ForEach-Object { $_.Trim() }

# Get all printers and ports locally
$allPrinters = Get-Printer
$allPorts = Get-PrinterPort | Select-Object Name, PrinterHostAddress

# Initialize the log file with headers
"Printer Name,Original Comment,Updated Comment" | Set-Content -Path $LogFile

foreach ($printerName in $PrinterNames) {
    # Extract the planning unit from the printer name (assumes it's the first part before '-')
    $planningUnit = $printerName -split "-" | Select-Object -First 1

    # Find the printer object
    $printer = $allPrinters | Where-Object { $_.Name -eq $printerName }
    if (-Not $printer) {
        Write-Host "Warning: Printer '$printerName' not found on this server. Skipping..."
        continue
    }

    # Get the current comment for the printer
    $currentComment = $printer.Comment
    if (-Not $currentComment) {
        $currentComment = "No current comment"
    }

    # Find the associated port and get the correct IP/hostname
    $port = $allPorts | Where-Object { $_.Name -eq $printer.PortName }
    $hostOrIP = if ($port) { $port.PrinterHostAddress } else { "Unknown" }

    # Resolve to IP if it's a hostname
    if ($hostOrIP -match "[a-zA-Z]") {
        try {
            $resolvedIP = (Resolve-DnsName -Name $hostOrIP -ErrorAction Stop).IPAddress
        } catch {
            Write-Host "Error resolving hostname '$hostOrIP': $_"
            $resolvedIP = "Unknown IP"
        }
    } else {
        $resolvedIP = $hostOrIP  # It's already an IP
    }

    # Clean the printer model name by removing "AltaLink", "VersaLink", "PCL6", and "Copier-Printer"
    $cleanModel = if ($printer.DriverName) {
        $printer.DriverName -replace "AltaLink|VersaLink|PCL6|Copier-Printer", "" -replace "\s+", " "
    } else {
        "Unknown Model"
    }

    # Construct the new comment in the format: "PlanningUnit - Model - IP Address"
    $comment = "$planningUnit - $cleanModel - $resolvedIP".Trim()

    # Format the log entry correctly and remove any newlines from comments
    $logEntry = '"{0}","{1}","{2}"' -f $printerName, ($currentComment -replace '[\r\n]+', ' ' -replace '"', '""'), ($comment -replace '"', '""')

    # Append the entry to the log file
    Add-Content -Path $LogFile -Value $logEntry

    # Display current comment and the new comment to be applied
    Write-Host "Current Comment for '$printerName': '$currentComment'"
    Write-Host "Updating '$printerName' with new comment: '$comment'"

    # Apply the change
    Set-Printer -Name $printerName -Comment $comment -WhatIf
}

Write-Host "Changes have been logged to: $LogFile"
