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
