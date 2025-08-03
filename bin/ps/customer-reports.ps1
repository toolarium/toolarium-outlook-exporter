###############################################################################
#
# customer-reports.ps1
#
# Copyright by toolarium, all rights reserved.
#
# This file is part of the toolarium outlook-exporter.
#
# The outlook-exporter is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# The outlook-exporter is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with Foobar. If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################


###############################################################################
# Settings
###############################################################################
$LINE="----------------------------------------------------------------------------------------"
$userFormatStr="dd/MM/yyyy"
$dataPathName = "data"
$customerPathName = "customer"
$reportPathName = "reports"
$currentPath = Get-Location
$customerInputPath = "$currentPath\$dataPathName\$customerPathName"
$reportOutputPath = "$currentPath\$dataPathName\$reportPathName"


###############################################################################
# Get customer files
###############################################################################
function Get-CustomerFiles {
    param (
        [string]$CustomerPath
    )
    
    $result = @{}
    if (-not (Test-Path -Path $CustomerPath)) {
        Write-Host ".: WARNING: Customer path not found: $CustomerPath"
        return $result
    }
    
    # Get all CSV files in customer directory
    $csvFiles = Get-ChildItem -Path $CustomerPath -Filter "*.csv"
    
    # Group files by customer name
    foreach ($file in $csvFiles) {
        $fileName = $file.BaseName
        
        # Parse filename pattern: YYYYMM-type-export-CustomerName
        if ($fileName -match "^(\d{6})-(calendar|mail)-export-(.+)$") {
            $period = $matches[1]
            $type = $matches[2]
            $customer = $matches[3]
            
            $key = "$period-$customer"
            
            if (-not $result.ContainsKey($key)) {
                $result[$key] = @{
                    Period = $period
                    Customer = $customer
                    Calendar = $null
                    Mail = $null
                }
            }
            
            if ($type -eq "calendar") {
                $result[$key].Calendar = $file.FullName
            } elseif ($type -eq "mail") {
                $result[$key].Mail = $file.FullName
            }
        }
    }
    
    return $result
}

function Convert-ToTimeTrackingEntry {
    param (
        [object]$SourceRow,
        [string]$Type  # "Calendar" or "Email"
    )
    
    # Parse the date to ensure proper sorting
    try {
        $parsedDate = [DateTime]::ParseExact($SourceRow.Date, $userFormatStr, $null)
    } catch {
        # Fallback parsing
        try {
            $parsedDate = [DateTime]::Parse($SourceRow.Date)
        } catch {
            Write-Host ".: WARNING: Could not parse date '$($SourceRow.Date)'"
            $parsedDate = Get-Date
        }
    }
    
    # Build description based on type
    $description = ""
    if ($Type -eq "Calendar") {
        $description = "Meeting: $($SourceRow.Subject)"
        if ($SourceRow.Location -and $SourceRow.Location.Trim() -ne "") {
            #$description += " (Location: $($SourceRow.Location))"
        }
        if ($SourceRow.Start -and $SourceRow.End -and $SourceRow.Start.Trim() -ne "" -and $SourceRow.End.Trim() -ne "") {
            #$description += " [$($SourceRow.Start)-$($SourceRow.End)]"
        }
    } else {
        $description = "Exchange: $($SourceRow.Subject)"
        if ($SourceRow.Start -and $SourceRow.Start.Trim() -ne "") {
            #$description += " [Sent: $($SourceRow.Start)]"
        }
    }
    
    # Get person name (try CalendarOwner first, then fallback)
    $person = ""
    if ($SourceRow.CalendarOwner -and $SourceRow.CalendarOwner.Trim() -ne "") {
        $person = $SourceRow.CalendarOwner
    } elseif ($SourceRow.OwnerEmail -and $SourceRow.OwnerEmail.Trim() -ne "") {
        $person = $SourceRow.OwnerEmail
    } else {
        $person = "Unknown"
    }
    
    # Get hours (try Duration first, then fallback to 0)
    $hours = 0
    if ($SourceRow.Duration -and $SourceRow.Duration -ne "") {
        try {
            $hours = [double]$SourceRow.Duration
        } catch {
            $hours = 0
        }
    }
    
    return [PSCustomObject]@{
        Date = $SourceRow.Date
        SortDate = $parsedDate
        Person = $person
        Hours = $hours
        Description = $description
        Type = $Type
    }
}

function Combine-CustomerData {
    param (
        [string]$CalendarFile,
        [string]$MailFile,
        [string]$Customer,
        [string]$Period
    )
    
    $combinedData = @()
    
    # Process calendar file
    if ($CalendarFile -and (Test-Path -Path $CalendarFile)) {
        try {
            $calendarData = Import-Csv -Path $CalendarFile
            #Write-Host "   - Calendar: $($calendarData.Count) entries"
            
            foreach ($row in $calendarData) {
                $combinedData += Convert-ToTimeTrackingEntry -SourceRow $row -Type "Calendar"
            }
        } catch {
            Write-Host "   - ERROR reading calendar file: $_"
        }
    } else {
        #Write-Host "   - Calendar: No file found"
    }
    
    # Process mail file
    if ($MailFile -and (Test-Path -Path $MailFile)) {
        try {
            $mailData = Import-Csv -Path $MailFile
            #Write-Host "   - Email: $($mailData.Count) entries"
            
            foreach ($row in $mailData) {
                $combinedData += Convert-ToTimeTrackingEntry -SourceRow $row -Type "Email"
            }
        } catch {
            Write-Host "   - ERROR reading mail file: $_"
        }
    } else {
        #Write-Host "   - Email: No file found"
    }
    
    return $combinedData
}


###############################################################################
# Main Processing
###############################################################################
Write-Host $LINE
Write-Host "   Combining calendar and email data for customers..."
Write-Host $LINE

if (-not (Test-Path -Path $reportOutputPath)) {
    New-Item -Path $reportOutputPath -ItemType Directory | Out-Null
    #Write-Host ".: Created combined output directory: $reportOutputPath"
}

# Get all customer files grouped by customer and period
$customerFiles = Get-CustomerFiles -CustomerPath $customerInputPath
Write-Host ".: Found $($customerFiles.Count) customer/period reports to process"

if ($customerFiles.Count -eq 0) {
    Write-Host ".: No customer files found in: $customerInputPath"
    Write-Host ".: Make sure you have run the customer filter script first"
    Exit 0
}

$totalReport = 0
$totalHours = 0

foreach ($key in $customerFiles.Keys) {
    $customerInfo = $customerFiles[$key]
    $customer = $customerInfo.Customer
    $period = $customerInfo.Period
    $calendarFile = $customerInfo.Calendar
    $mailFile = $customerInfo.Mail
    
    Write-Host ".: Processing $customer ($period)..."
    
    # Combine data from both sources
    $combinedData = Combine-CustomerData -CalendarFile $calendarFile -MailFile $mailFile -Customer $customer -Period $period
    
    if ($combinedData.Count -gt 0) {
        # Sort by date and time
        $sortedData = $combinedData | Sort-Object SortDate
        
        # Calculate total hours for this customer
        $customerHours = ($sortedData | Measure-Object -Property Hours -Sum).Sum
        $totalHours += $customerHours
        
        # Create final export data (remove helper columns)
        $exportData = @()
        foreach ($entry in $sortedData) {
            $exportData += [PSCustomObject]@{
                Date = $entry.Date
                Person = $entry.Person
                Hours = $entry.Hours
                Description = $entry.Description
            }
        }
        
        # Generate output filename
        $outputFileName = "$period-$customer.csv"
        $outputPath = "$reportOutputPath\$outputFileName"
        
        # Export to CSV
        $exportData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        
        Write-Host "   - Reported: $($combinedData.Count) total entries ($customerHours hours)"
        #Write-Host "   - Exported to: $outputFileName"
        $totalReport++
    } else {
        Write-Host "   - No data found for $customer ($period)"
    }
    
    # Clean up
    $combinedData = $null
    $sortedData = $null
    $exportData = $null
}

# Final summary
#Write-Host ".: Created $totalReport report files"
Write-Host ".: Total hours across all customers: $totalHours"
Write-Host ""
Write-Host $LINE
Write-Host ".: Output directory: $reportOutputPath"
Write-Host $LINE

# Show sample of what was created
if ($totalReport -gt 0) {
    #Write-Host ""
    #Write-Host ".: Sample combined file structure:"
    #$sampleFiles = Get-ChildItem -Path $reportOutputPath -Filter "*.csv" | Select-Object -First 3
    #foreach ($file in $sampleFiles) {
    #    Write-Host "   - $($file.Name)"
    #}
}