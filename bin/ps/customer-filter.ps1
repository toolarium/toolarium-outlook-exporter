###############################################################################
#
# customer-filter.ps1
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
$caseSensitive = $false  # Set to $false for case-insensitive filtering
$excludePrivateWithoutLocation = $true  # Set to $false to include private entries without location
$dataPathName = "data"
$configPathName = "config"
$customerPathName = "customer"
$currentPath = Get-Location
$dataInputPath = "$currentPath\$dataPathName"
$dataOutputPath = "$currentPath\$dataPathName\$customerPathName"
$configPath = "$currentPath\$configPathName"
$customerFilterConfigPath = "$currentPath\$configPathName\$customerPathName-filter"
$printLoadFilter=$true


###############################################################################
# Filter words
###############################################################################
function Get-FilterWords {
    param (
        [string]$RelativePath
    )
    
    $result = @()
    if (Test-Path -Path $RelativePath) {
        try {
            $rawContent = Get-Content -Path $RelativePath -ErrorAction Stop
			if ($printLoadFilter) {
				Write-Host "   - Loaded filter file:" $RelativePath
			}
            
            # Process each line and trim whitespace, but keep internal spaces
            foreach ($line in $rawContent) {
                $trimmedLine = $line.Trim()
                if ($trimmedLine -ne "" -and -not $trimmedLine.StartsWith("#")) {
                    $result += $trimmedLine
                }
            }
        } catch {
            Write-Host "   - Error reading filter file:" $RelativePath
            $result = @()
        }
    } else {
															  
        $result = @()
    }
    
    return $result
}


###############################################################################
# Get customer duration settings
###############################################################################
function Get-CustomerDurationSettings {
    param (
        [string]$CustomerFilterConfigPath,
        [string]$CustomerName
    )
    
    # Default values if no config file found
    $settings = @{
        DefaultEmailDuration = 1.0
        AdditionalEmailDuration = 0.5
    }
    
    $durationFilePath = "$CustomerFilterConfigPath\$CustomerName-duration.txt"
    
    if (Test-Path -Path $durationFilePath) {
        try {
            $lines = Get-Content -Path $durationFilePath -ErrorAction Stop
            Write-Host "   - Loaded duration settings: $durationFilePath"
            
            foreach ($line in $lines) {
                $trimmedLine = $line.Trim()
                if ($trimmedLine -ne "" -and -not $trimmedLine.StartsWith("#")) {
                    if ($trimmedLine.Contains("=")) {
                        # Format: KEY=VALUE
                        $parts = $trimmedLine.Split("=", 2)
                        $key = $parts[0].Trim()
                        $value = $parts[1].Trim()
                        
                        try {
                            $numericValue = [double]$value
                            switch ($key.ToUpper()) {
                                "DEFAULT_EMAIL_DURATION" { $settings.DefaultEmailDuration = $numericValue }
                                "ADDITIONAL_EMAIL_DURATION" { $settings.AdditionalEmailDuration = $numericValue }
                                "DEFAULT" { $settings.DefaultEmailDuration = $numericValue }
                                "ADDITIONAL" { $settings.AdditionalEmailDuration = $numericValue }
                            }
                        } catch {
                            Write-Host "   - WARNING: Invalid numeric value '$value' for '$key'"
                        }
                    } else {
                        # Simple format: two numbers on separate lines
                        try {
                            $numericValue = [double]$trimmedLine
                            if ($settings.DefaultEmailDuration -eq 1.0) {
                                $settings.DefaultEmailDuration = $numericValue
                            } else {
                                $settings.AdditionalEmailDuration = $numericValue
                            }
                        } catch {
                            Write-Host "   - WARNING: Invalid numeric value '$trimmedLine'"
                        }
                    }
                }
            }
        } catch {
            Write-Host "   - Error reading duration file: $durationFilePath"
        }
    } else {
        #Write-Host "   - No duration config found, using defaults (Default: $($settings.DefaultEmailDuration)h, Additional: $($settings.AdditionalEmailDuration)h)"
    }
    
    return $settings
}


###############################################################################
# Calculate customer-specific email duration
###############################################################################
function Calculate-CustomerEmailDuration {
    param (
        [array]$CustomerEmails,
        [hashtable]$DurationSettings,
        [string]$CustomerName
    )
    
    if ($CustomerEmails.Count -eq 0) {
        return $CustomerEmails
    }
    
    Write-Host "   - Calculating email durations for $CustomerName (Default: $($DurationSettings.DefaultEmailDuration)h, Additional: $($DurationSettings.AdditionalEmailDuration)h)"
    
    # Group emails by date and cleaned subject
    $emailGroups = @{}
    
    foreach ($email in $CustomerEmails) {
        $dateKey = $email.Date
        $subjectKey = if ($email.Subject) { $email.Subject.ToString().ToLower() } else { "no-subject" }
        $groupKey = "$dateKey|$subjectKey"
        
        if (-not $emailGroups.ContainsKey($groupKey)) {
            $emailGroups[$groupKey] = @()
        }
        $emailGroups[$groupKey] += $email
    }
    
    # Calculate durations based on customer settings and occurrence count
    $adjustedEmails = 0
    foreach ($groupKey in $emailGroups.Keys) {
        $emailsInGroup = $emailGroups[$groupKey]
        
        if ($emailsInGroup.Count -gt 1) {
            # Multiple emails for same subject on same day
            $totalDuration = 0
            
            for ($i = 0; $i -lt $emailsInGroup.Count; $i++) {
                if ($i -eq 0) {
                    # First occurrence
                    $emailsInGroup[$i].Duration = $DurationSettings.DefaultEmailDuration
                    $totalDuration += $DurationSettings.DefaultEmailDuration
                } elseif ($i -eq 1) {
                    # Second occurrence - keep default
                    $emailsInGroup[$i].Duration = $DurationSettings.DefaultEmailDuration
                    $totalDuration += $DurationSettings.DefaultEmailDuration
                } else {
                    # Third and beyond - add additional time
                    $additionalDuration = $DurationSettings.DefaultEmailDuration + (($i - 1) * $DurationSettings.AdditionalEmailDuration)
                    $emailsInGroup[$i].Duration = $additionalDuration
                    $totalDuration += $additionalDuration
                    $adjustedEmails++
                }
            }
            
            # Debug output for groups with multiple emails
            $parts = $groupKey.Split('|')
            $dateKey = $parts[0]
            $subjectKey = $parts[1]
            Write-Host "     * $dateKey '$subjectKey': $($emailsInGroup.Count) emails → total: $totalDuration hours"
        } else {
            # Single email - use default duration
            $emailsInGroup[0].Duration = $DurationSettings.DefaultEmailDuration
        }
    }
    
    if ($adjustedEmails -gt 0) {
        Write-Host "   - Applied duration adjustments to $adjustedEmails emails (3rd+ occurrences)"
    }
    
    return $CustomerEmails
}


###############################################################################
# Clean email subject from RE:, FW:, FWD: prefixes
###############################################################################
function Clean-EmailSubject {
    param (
        [string]$Subject
    )
    
    if (-not $Subject -or $Subject.Trim() -eq "") {
        return $Subject
    }
    
    $cleanSubject = $Subject.Trim()
    
    # List of prefixes to remove (case insensitive)
    $prefixesToRemove = @('RE:', 'FW:', 'FWD:', 'AW:', 'WG:', 'RE :', 'FW :', 'FWD :', 'AW :', 'WG :')
    
    # Keep removing prefixes until none are found (handles nested prefixes like "RE: FW: RE:")
    $maxIterations = 10  # Prevent infinite loops
    $iteration = 0
    
    do {
        $originalSubject = $cleanSubject
        $iteration++
        
        # Try each prefix
        foreach ($prefix in $prefixesToRemove) {
            # Check if subject starts with this prefix (case insensitive)
            if ($cleanSubject.ToLower().StartsWith($prefix.ToLower())) {
                $cleanSubject = $cleanSubject.Substring($prefix.Length).Trim()
                break  # Exit foreach loop if we found a match
            }
        }
        
    } while ($cleanSubject -ne $originalSubject -and $cleanSubject -ne "" -and $iteration -lt $maxIterations)
    
    # Return cleaned subject or original if cleaning resulted in empty string
    if ($cleanSubject -ne "" -and $cleanSubject.Trim() -ne "") {
        return $cleanSubject.Trim()
    } else {
        return $Subject
    }
}


###############################################################################
# Test function for email subject cleaning (optional - for debugging)
###############################################################################
function Test-EmailSubjectCleaning {
    Write-Host ".: Testing email subject cleaning:"
    
    $testSubjects = @(
        "RE: Support Day",
        "FW: Project Update", 
        "RE: FW: Meeting Tomorrow",
        "FWD: ClientA Contract",
        "RE: RE: Important Notice",
        "AW: German Reply",
        "Normal Subject Without Prefix",
        "RE:No Space After Colon",
        "RE : Space Before Colon",
        ""
    )
    
    foreach ($subject in $testSubjects) {
        $cleaned = Clean-EmailSubject -Subject $subject
        Write-Host "   - '$subject' → '$cleaned'"
    }
    Write-Host ""
}

# Uncomment the next line to test the cleaning function:
# Test-EmailSubjectCleaning


###############################################################################
# Main Processing
###############################################################################
Write-Host $LINE
Write-Host "  Processing calendar and email data for customers..."
Write-Host $LINE

if (-not (Test-Path -Path $dataOutputPath)) {
    New-Item -Path $dataOutputPath -ItemType Directory | Out-Null
}

# Read all data source files (CSV files from calendar export)
$sourceFiles = Get-ChildItem -Path $dataInputPath -Filter "*.csv"
Write-Host ".: Found $($sourceFiles.Count) source files to process"

foreach ($sourceFile in $sourceFiles) {
    Write-Host ".: Processing source file: $($sourceFile.Name)"
    
    # Import the CSV data
    $csv = $null
    try {
        $csv = Import-Csv -Path $sourceFile.FullName
    } catch {
        Write-Host ".: ERROR: Could not read $($sourceFile.Name): $_"
        continue
    }
    
    if ($csv.Count -eq 0) {
        Write-Host ".: No data in $($sourceFile.Name) - skipping"
        $csv = $null
        continue
    }
    
    # Filter out private entries with no location before processing customers
    $originalCount = $csv.Count
    $filteredCsv = @()
    $excludedCount = 0
    $excludedEntries = @()
    
    if ($excludePrivateWithoutLocation) {
        Write-Host ".: Filtering out private entries with no location..."
        $hasPrivateColumn = $csv[0].PSObject.Properties.Name -contains "Private"
        $hasLocationColumn = $csv[0].PSObject.Properties.Name -contains "Location"
        
        if (-not $hasPrivateColumn) {
            Write-Host ".: WARNING: No 'Private' column found - private filtering disabled"
        }
        if (-not $hasLocationColumn) {
            Write-Host ".: WARNING: No 'Location' column found - location filtering disabled"
        }
        
        foreach ($row in $csv) {
            $shouldExclude = $false
            
            # Check if entry should be excluded (Private = True AND Location is empty)
            if ($hasPrivateColumn -and $hasLocationColumn) {
                $isPrivate = $false
                $hasLocation = $true
                
                # Check if Private (handle different possible values)
                if ($row.Private) {
                    $privateValue = $row.Private.ToString().ToLower()
                    $isPrivate = ($privateValue -eq "true" -or $privateValue -eq "yes" -or $privateValue -eq "1")
                }
                
                # Check if Location is empty
                if (-not $row.Location -or $row.Location.ToString().Trim() -eq "") {
                    $hasLocation = $false
                }
                
                # Exclude if Private AND no Location
                if ($isPrivate -and -not $hasLocation) {
                    $shouldExclude = $true
                    $excludedCount++
                    $excludedEntries += "   - $($row.Date) $($row.Start): $($row.Subject)"
                }
            }
            
            # Add to filtered CSV if not excluded
            if (-not $shouldExclude) {
                $filteredCsv += $row
            }
        }
        
        # Update CSV reference to use filtered data
        $csv = $filteredCsv
        
        Write-Host ".: Filtered: $originalCount total -> $($csv.Count) for processing ($excludedCount excluded as private/no location)"
        
        # Optionally show excluded entries (uncomment if needed)
        #if ($excludedCount -gt 0 -and $excludedCount -le 10) {
        #    Write-Host ".: Excluded entries:"
        #    $excludedEntries | ForEach-Object { Write-Host $_ }
        #} elseif ($excludedCount -gt 10) {
        #    Write-Host ".: Excluded entries (showing first 10):"
        #    $excludedEntries[0..9] | ForEach-Object { Write-Host $_ }
        #    Write-Host "   - ... and $($excludedCount - 10) more"
        #}
    } else {
        Write-Host ".: Private entry filtering disabled - processing all entries"
        $filteredCsv = $csv
    }
    
    if ($csv.Count -eq 0) {
        Write-Host ".: No appointments left after filtering - skipping"
        $csv = $null
        $filteredCsv = $null
        continue
    }
    
    # Clean email subjects from all rows (for both calendar and email data)
    $hasSubjectColumn = $csv[0].PSObject.Properties.Name -contains "Subject"
    if ($hasSubjectColumn) {
        $cleanedSubjects = 0
        $newCsv = @()
        #Write-Host ".: Cleaning email subject prefixes..."
        
        foreach ($row in $csv) {
            # Create a completely new object to ensure the Subject property can be modified
            $newRow = [PSCustomObject]@{}
            
            # Copy all properties from original row
            foreach ($property in $row.PSObject.Properties) {
                if ($property.Name -eq "Subject" -and $property.Value -and $property.Value.ToString().Trim() -ne "") {
                    $originalSubject = $property.Value.ToString()
                    $cleanedSubject = Clean-EmailSubject -Subject $originalSubject
                    
                    if ($cleanedSubject -ne $originalSubject) {
                        #Write-Host "   - Cleaning: '$originalSubject' → '$cleanedSubject'"
                        $cleanedSubjects++
                    }
                    
                    # Add the cleaned subject to the new object
                    $newRow | Add-Member -MemberType NoteProperty -Name $property.Name -Value $cleanedSubject
                } else {
                    # Copy other properties as-is
                    $newRow | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                }
            }
            
            $newCsv += $newRow
        }
        
        # Replace the original CSV array with the cleaned version
        $csv = $newCsv
        
        if ($cleanedSubjects -gt 0) {
            #Write-Host ".: Cleaned $cleanedSubjects email subject prefixes (RE:, FW:, etc.)"
        } else {
            #Write-Host ".: No email subject prefixes found to clean"
        }
    }
    
    # Track which rows have been matched by any customer
    $matchedRowIndices = @()
    
    # Read all customer config files
    $configFiles = Get-ChildItem -Path $customerFilterConfigPath -Filter "*.txt"
    #Write-Host ".: Found $($configFiles.Count) customer config files"
    
    foreach ($configFile in $configFiles) {
        #Write-Host ".: Processing customer config: $($configFile.Name)"
        
        # Read customer keywords from config file
        $customerKeyWordsFilter = Get-FilterWords -RelativePath $configFile.FullName
        
        # Read customer attendees from attendee file (same name but with -attendees suffix)
        $customerName = $configFile.BaseName
        $attendeeFileName = "$customerName-attendees.txt"
        $attendeeFilePath = "$customerFilterConfigPath\$attendeeFileName"
        $customerAttendeesFilter = Get-FilterWords -RelativePath $attendeeFilePath
        
        # Read customer duration settings for email processing
        $customerDurationSettings = Get-CustomerDurationSettings -CustomerFilterConfigPath $customerFilterConfigPath -CustomerName $customerName

        # Skip if no keywords AND no attendees found
        if ($customerKeyWordsFilter.Count -eq 0 -and $customerAttendeesFilter.Count -eq 0) {
            #Write-Host ".: No keywords or attendees found for $customerName - skipping"
            $customerKeyWordsFilter = $null
            $customerAttendeesFilter = $null
            continue
        }
        
        # Apply case sensitivity setting to keywords
        if (-not $caseSensitive -and $customerKeyWordsFilter.Count -gt 0) {
            $customerKeyWordsFilter = $customerKeyWordsFilter | ForEach-Object { $_.ToLower() }
        }
        
        # Apply case sensitivity setting to attendees
        if (-not $caseSensitive -and $customerAttendeesFilter.Count -gt 0) {
            $customerAttendeesFilter = $customerAttendeesFilter | ForEach-Object { $_.ToLower() }
        }
        
        # Debug: Show keywords being used (uncomment for troubleshooting)
        if ($customerKeyWordsFilter.Count -gt 0) {
            #Write-Host "   - Keywords for $customerName`: $($customerKeyWordsFilter -join ', ')"
        }
        if ($customerAttendeesFilter.Count -gt 0) {
            #Write-Host "   - Attendees for $customerName`: $($customerAttendeesFilter -join ', ')"
        }
        
        #Write-Host ".: $customerName - Keywords: $($customerKeyWordsFilter.Count), Attendees: $($customerAttendeesFilter.Count)"
        
        # Check if required columns exist (Subject was already checked during cleaning)
        $hasAttendeesColumn = $csv[0].PSObject.Properties.Name -contains "Attendees"
        
        if (-not $hasSubjectColumn -and $customerKeyWordsFilter.Count -gt 0) {
            Write-Host ".: WARNING: No 'Subject' column found but keywords defined - subject filtering disabled"
        }
        if (-not $hasAttendeesColumn -and $customerAttendeesFilter.Count -gt 0) {
            Write-Host ".: WARNING: No 'Attendees' column found but attendees defined - attendee filtering disabled"
        }
        
        $matchedRows = @()
        
        # Filter on Subject keywords AND/OR Attendees
        for ($i = 0; $i -lt $csv.Count; $i++) {
            $row = $csv[$i]
            $matchFound = $false
            
            # Check Subject keywords (if available)
            if ($hasSubjectColumn -and $customerKeyWordsFilter.Count -gt 0 -and -not $matchFound) {
                $subjectValue = $row.Subject
                if ($subjectValue -and $subjectValue.ToString().Trim() -ne "") {
                    $subjectValueToCheck = if ($caseSensitive) { $subjectValue.ToString() } else { $subjectValue.ToString().ToLower() }
                    
                    foreach ($keyWord in $customerKeyWordsFilter) {
                        # Ensure keyword is properly trimmed but preserve internal spaces
                        $cleanKeyword = $keyWord.Trim()
                        if ($cleanKeyword -ne "" -and $subjectValueToCheck.Contains($cleanKeyword)) {
                            $matchedRows += $row
                            if ($matchedRowIndices -notcontains $i) {
                                $matchedRowIndices += $i
                            }
                            $matchFound = $true
                            # Debug output (uncomment for troubleshooting)
                            # Write-Host ".: SUBJECT MATCH: '$subjectValue' contains '$cleanKeyword'"
                            break
                        }
                    }
                }
            }
            
            # Check Attendees (if available and no subject match found)
            if ($hasAttendeesColumn -and $customerAttendeesFilter.Count -gt 0 -and -not $matchFound) {
                $attendeesValue = $row.Attendees
                if ($attendeesValue -and $attendeesValue.ToString().Trim() -ne "") {
                    $attendeesValueToCheck = if ($caseSensitive) { $attendeesValue.ToString() } else { $attendeesValue.ToString().ToLower() }
                    
                    foreach ($attendeeFilter in $customerAttendeesFilter) {
                        # Ensure attendee filter is properly trimmed but preserve internal spaces
                        $cleanAttendeeFilter = $attendeeFilter.Trim()
                        if ($cleanAttendeeFilter -ne "" -and $attendeesValueToCheck.Contains($cleanAttendeeFilter)) {
                            $matchedRows += $row
                            if ($matchedRowIndices -notcontains $i) {
                                $matchedRowIndices += $i
                            }
                            $matchFound = $true
                            # Debug output (uncomment for troubleshooting)
                            # Write-Host ".: ATTENDEE MATCH: '$attendeesValue' contains '$cleanAttendeeFilter'"
                            break
                        }
                    }
                }
            }
        }
        
        # Apply customer-specific duration calculation for email files
        $isEmailFile = $sourceFile.Name.ToLower().Contains("mail-export")
        if ($isEmailFile -and $matchedRows.Count -gt 0) {
            Write-Host ".: Applying customer-specific email duration settings for $customerName"
            $matchedRows = Calculate-CustomerEmailDuration -CustomerEmails $matchedRows -DurationSettings $customerDurationSettings -CustomerName $customerName
        }
		
        #Write-Host ".: Found $($matchedRows.Count) matching rows for $($configFile.BaseName)"
        
        # Export matched rows if any found
        if ($matchedRows.Count -gt 0) {
            $customerName = $configFile.BaseName
            $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile.Name)
            $outputFileName = "$sourceFileName-$customerName.csv"
            $outputCsvPath = "$dataOutputPath\$outputFileName"
            
            $matchedRows | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8
            #Write-Host ".: Exported $($matchedRows.Count) entries to: $outputFileName"
            
            # Show total hours for this customer if it's an email file
            if ($isEmailFile) {
                $totalHours = ($matchedRows | Measure-Object -Property Duration -Sum).Sum
                #Write-Host ".: Total email time for $customerName`: $totalHours hours"
            }
			
        } else {
            #Write-Host ".: No matching entries found for $($configFile.BaseName)"
        }
        
        # Clean up only the processed customer data
        $matchedRows = $null
        $customerKeyWordsFilter = $null
        $customerAttendeesFilter = $null
        $customerDurationSettings = $null
    }
    
    # Create file with unmatched entries
    $unmatchedRows = @()
    for ($i = 0; $i -lt $csv.Count; $i++) {
        if ($matchedRowIndices -notcontains $i) {
            $unmatchedRows += $csv[$i]
        }
    }
    
    if ($unmatchedRows.Count -gt 0) {
        $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile.Name)
        $unmatchedFileName = "$sourceFileName-Unmatched.csv"
        $unmatchedCsvPath = "$dataOutputPath\$unmatchedFileName"
        
        $unmatchedRows | Export-Csv -Path $unmatchedCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host ".: Exported $($unmatchedRows.Count) unmatched entries to: $unmatchedFileName"
		$printLoadFilter=$false
    } else {
        Write-Host ".: All entries were matched to customers - no unmatched file created"
    }
    
    # Clean up after processing all customers for this source file
    $csv = $null
    $configFiles = $null
    $matchedRowIndices = $null
    $unmatchedRows = $null
    $filteredCsv = $null
    $excludedEntries = $null
    $newCsv = $null	
    
    # Force garbage collection for large datasets
    if ($sourceFiles.Count -gt 5 -or (Get-Process -Name "powershell*" | Measure-Object WorkingSet -Sum).Sum -gt 500MB) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Final cleanup
$sourceFiles = $null

# Final garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
