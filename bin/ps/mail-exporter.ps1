###############################################################################
#
# mail-exportter.ps1
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
$msFormatStr="dd/MM/yyyy"
$userFormatStr="dd/MM/yyyy"
$dataPathName = "data"
$configPathName = "config"

# Email duration constants
$DEFAULT_EMAIL_DURATION = 0.5     # Default hours for first email
$ADDITIONAL_EMAIL_DURATION = 0.5  # Additional hours for 2nd+ emails per day/subject


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
# Calculate email duration based on daily subject occurrences
###############################################################################
function Calculate-EmailDuration {
    param (
        [array]$AllEmails
    )
    
    #    Write-Host ".: Calculating email durations based on daily subject frequency..."

    
    # Group emails by date and cleaned subject
    $emailGroups = @{}
    
    foreach ($email in $AllEmails) {
        $dateKey = $email.Date
        $subjectKey = $email.CleanedSubject.ToLower() # Case insensitive grouping
        $groupKey = "$dateKey|$subjectKey"
        
        if (-not $emailGroups.ContainsKey($groupKey)) {
            $emailGroups[$groupKey] = @()
        }
        $emailGroups[$groupKey] += $email
    }
    
    # Calculate durations based on occurrence count
    $totalAdjustments = 0
    foreach ($groupKey in $emailGroups.Keys) {
        $emailsInGroup = $emailGroups[$groupKey]
        $occurrenceCount = $emailsInGroup.Count
        
        # Apply duration logic:
        # 1st occurrence: 0.5 hours
        # 2nd+ occurrence: add 0.5 hours for each additional occurrence
        
        for ($i = 0; $i -lt $emailsInGroup.Count; $i++) {
            $email = $emailsInGroup[$i]
            
            if ($i -eq 0) {
                # First occurrence
                $email.Duration = $DEFAULT_EMAIL_DURATION
            #} elseif ($i -eq 1) {
            #    # Second occurrence - keep default
            #    $email.Duration = $DEFAULT_EMAIL_DURATION
            } else {
                # Third and beyond - add additional time
                $email.Duration = $DEFAULT_EMAIL_DURATION + (($i - 1) * $ADDITIONAL_EMAIL_DURATION)
                $totalAdjustments++
            }
        }
        
        # Debug output for groups with multiple emails
        if ($occurrenceCount -gt 1) {
            $parts = $groupKey.Split('|')
            $dateKey = $parts[0]
            $subjectKey = $parts[1]
			$totalDuration = ($emailsInGroup.Duration | Measure-Object -Sum).Sum
            #Write-Host "   - $dateKey '$subjectKey': $occurrenceCount emails → durations: $($emailsInGroup.Duration -join ', ') hours"
        }
    }
    
    if ($totalAdjustments -gt 0) {
        #Write-Host ".: Applied duration adjustments to $totalAdjustments emails (3rd+ occurrences)"
    }
    
    return $AllEmails
}


###############################################################################
# Get full attendee information with caching
###############################################################################
$global:attendeeCache = @{}
$global:cacheStats = @{
    Hits = 0
    Misses = 0
    Errors = 0
}

function Get-FullEmailRecipients {
    param($item)
    
    $recipientsList = @()
    try {
        if ($item.Recipients.Count -gt 0) {
            for ($i = 1; $i -le $item.Recipients.Count; $i++) {
                $recipient = $item.Recipients.Item($i)
                
                # Create cache key from recipient properties
                $cacheKey = ""
                try {
                    $cacheKey = $recipient.AddressEntry.ID # Use EntryID as primary cache key (most reliable)
                } catch {
                    $cacheKey = "$($recipient.Name)|$($recipient.Address)" # Fallback: use name + address combination
                }
                
                # Check cache first
                if ($global:attendeeCache.ContainsKey($cacheKey)) {
                    $recipientsList += $global:attendeeCache[$cacheKey]
                    $global:cacheStats.Hits++
                    continue
                }
                
                $global:cacheStats.Misses++ # Not in cache - resolve recipient info
                $resolvedRecipient = Resolve-SingleEmailRecipient -recipient $recipient
                
                $global:attendeeCache[$cacheKey] = $resolvedRecipient # Cache the result
                $recipientsList += $resolvedRecipient
            }
        }
    } catch {
        $global:cacheStats.Errors++ # Fallback to original method if Recipients collection fails
        return $item.To
    }
    return ($recipientsList -join "; ") # Join all recipients with semicolon
}

# Resolve single email recipient (separated for caching efficiency)
function Resolve-SingleEmailRecipient {
    param($recipient)
    
    try {
        $recipientName = $recipient.Name
        $recipientEmail = ""
        
        # Resolve email address based on type
        if ($recipient.AddressEntry.Type -eq "SMTP") {
            # Direct SMTP address - fastest
            $recipientEmail = $recipient.AddressEntry.Address
        } elseif ($recipient.AddressEntry.Type -eq "EX") {
            # Exchange address - slow, so caching is critical here
            try {
                $exchangeUser = $recipient.AddressEntry.GetExchangeUser()
                if ($exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
                    $recipientEmail = $exchangeUser.PrimarySmtpAddress
                } else {
                    $recipientEmail = $recipient.AddressEntry.Address
                }
            } catch {
                # Exchange resolution failed - use display address
                $recipientEmail = $recipient.AddressEntry.Address
            }
        } else {
            # Other address types
            $recipientEmail = $recipient.AddressEntry.Address
        }
        
        # Format as "Name <email>" or just email if name is same
        if ($recipientName -ne $recipientEmail -and $recipientEmail -ne "" -and $recipientName -ne "") {
            return "$recipientName <$recipientEmail>"
        } else {
            return if ($recipientEmail -ne "") { $recipientEmail } else { $recipientName }
        }
        
    } catch {
        # Ultimate fallback - just return the name
        return $recipient.Name
    }
}

function Clear-AttendeeCache {
    $global:attendeeCache.Clear()
    $global:cacheStats = @{ Hits = 0; Misses = 0; Errors = 0 }
    Write-Host ".: Attendee cache cleared"
}


###############################################################################
# Cache statistics and cleanup functions
###############################################################################
function Show-CacheStats {
    $total = $global:cacheStats.Hits + $global:cacheStats.Misses
    if ($total -gt 0) {
        $hitRate = [math]::Round(($global:cacheStats.Hits / $total) * 100, 1)
        Write-Host ".: Email Recipients Cache Stats:"
        Write-Host "   - Cache Size: $($global:attendeeCache.Count) entries"
        Write-Host "   - Cache Hits: $($global:cacheStats.Hits) ($hitRate%)"
        Write-Host "   - Cache Misses: $($global:cacheStats.Misses)"
        Write-Host "   - Errors: $($global:cacheStats.Errors)"
    }
}


###############################################################################
# Get Current Outlook User Information
###############################################################################
function Get-OutlookCurrentUser {
    param($namespace)
    
    $userInfo = @{
        DisplayName = ""
        EmailAddress = ""
        ExchangeName = ""
        AccountType = ""
        IsExchange = $false
    }
    
    try {
        # Method 1: Get current user from namespace
        $currentUser = $namespace.CurrentUser
        if ($currentUser) {
            $userInfo.DisplayName = $currentUser.Name
            
            try {
                # Try to get email address from AddressEntry
                if ($currentUser.AddressEntry) {
                    $addressEntry = $currentUser.AddressEntry
                    
                    if ($addressEntry.Type -eq "EX") {
                        # Exchange user - get detailed info
                        $userInfo.IsExchange = $true
                        $userInfo.AccountType = "Exchange"
                        $userInfo.ExchangeName = $addressEntry.Address
                        
                        try {
                            $exchangeUser = $addressEntry.GetExchangeUser()
                            if ($exchangeUser) {
                                $userInfo.EmailAddress = $exchangeUser.PrimarySmtpAddress
                                if ($exchangeUser.Name) {
                                    $userInfo.DisplayName = $exchangeUser.Name
                                }
                            }
                        } catch {
                            Write-Host ".: Could not resolve Exchange user details"
                        }
                    } elseif ($addressEntry.Type -eq "SMTP") {
                        # Direct SMTP account
                        $userInfo.AccountType = "SMTP"
                        $userInfo.EmailAddress = $addressEntry.Address
                    } else {
                        $userInfo.AccountType = $addressEntry.Type
                        $userInfo.EmailAddress = $addressEntry.Address
                    }
                }
            } catch {
                Write-Host ".: Could not access AddressEntry for current user"
            }
        }
        
        # Method 2: Alternative - get from calendar folder properties
        if ($userInfo.EmailAddress -eq "" -or $userInfo.DisplayName -eq "") {
            try {
                $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
                
                # Try to get owner information from folder
                if ($calendar.FolderPath) {
                    $userInfo.ExchangeName = $calendar.FolderPath
                }
                
                # Get from session if available
                if ($namespace.Session.CurrentUser) {
                    if ($userInfo.DisplayName -eq "") {
                        $userInfo.DisplayName = $namespace.Session.CurrentUser.Name
                    }
                }
            } catch {
                Write-Host ".: Could not access calendar folder properties"
            }
        }
        
        # Method 3: Get from Accounts collection (Outlook 2010+)
        try {
            if ($namespace.Session.Accounts.Count -gt 0) {
                $primaryAccount = $namespace.Session.Accounts.Item(1)
                if ($userInfo.EmailAddress -eq "" -and $primaryAccount.SmtpAddress) {
                    $userInfo.EmailAddress = $primaryAccount.SmtpAddress
                }
                if ($userInfo.DisplayName -eq "" -and $primaryAccount.DisplayName) {
                    $userInfo.DisplayName = $primaryAccount.DisplayName
                }
                if ($userInfo.AccountType -eq "" -and $primaryAccount.AccountType) {
                    $userInfo.AccountType = "Account: " + $primaryAccount.AccountType
                }
            }
        } catch {
            Write-Host ".: Could not access Accounts collection"
        }
        
    } catch {
        Write-Host ".: Error getting current user information: $_"
    }
    
    return $userInfo
}


###############################################################################
# Display user information nicely
###############################################################################
function Show-CurrentUserInfo {
    param($userInfo)
    
    Write-Host ".: Current Outlook User Information:"
    Write-Host "   - Display Name: $($userInfo.DisplayName)"
    Write-Host "   - Email Address: $($userInfo.EmailAddress)"
    Write-Host "   - Account Type: $($userInfo.AccountType)"
    
    if ($userInfo.IsExchange) {
        Write-Host "   - Exchange Name: $($userInfo.ExchangeName)"
        Write-Host "   - Exchange Account: Yes"
    } else {
        Write-Host "   - Exchange Account: No"
    }
}


###############################################################################
# Initialize default date range (current month if no parameters)
###############################################################################
$currentDate = Get-Date
$startDate = Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1
$endDate = $startDate.AddMonths(1).AddDays(-1)

# Parse parameters
if ($args.count -gt 0) {
	$inputSplit = 0
	$charCount = 0
	$inputIsNumber=(($args[0] -is [string]) -and  ($args[0] -match "^\d+$"))
	if (-not $inputIsNumber) {
		$charCount = ($args[0].ToCharArray() | Where-Object {$_ -eq '.'} | Measure-Object).Count
		$inputIsDouble=(($args[0] -is [string]) -and  ($args[0] -match "^[\d\.\d]+$") -and $charCount -eq 1)
	}

	if ($args[0] -is [double]) {
	    $inputSplit = ($args[0].ToString('##.####')).split(".")
	} elseif ($inputIsDouble) {
	    $inputSplit = $args[0].split(".")
	}

    if (($args[0] -is [int]) -or ($args[0] -is [double]) -or $inputIsNumber -or $inputIsDouble) {
		$year = [int]::Parse((Get-Date).ToString("yyyy"))
		$month = [int]::Parse((Get-Date).ToString("MM"))
	
		if ($args[0] -is [int] -or $inputIsNumber) {
			if ($inputIsNumber) {
				$inputSplit = [int]::Parse($args[0])
			} else {
			    $inputSplit = $args
			}
		}
		
		if ($inputSplit.count -gt 0) {
			if(($inputSplit[0] -ge 1) -and (12 -ge $inputSplit[0])) { # then it's a month: 1 - 12			
				$month=[int]::Parse($inputSplit[0])
				
				if($inputSplit.count -gt 1) {
					$year = [int]::Parse($inputSplit[1])					
				}
			} else { # then it's a year
				$year=[int]::Parse($inputSplit[0])					
			}
			
			if(($month -ge 1) -and (12 -ge $month )) { # then it's a month: 1 - 12			
			} else {
				 Write-Host ".: ERROR: Invalid month input "$month" (1-12)!"
				 Exit 1
			}
			if ($year -ge 2000) {
			} else {
				 Write-Host ".: ERROR: Invalid year input "$year"!"
				 Exit 1
			}
		}
		
		$month = (Get-Date $year"-"$month"-01").ToString("MM")
		$firstDayOfMonth = Get-Date $year"-"$month"-01"
		$nextmonth = $firstDayOfMonth.AddMonths(1)
		$thefirst = Get-Date -Year $nextmonth.Year -Month $nextmonth.Month -Day 1
		$lastDayoOfMonth = $thefirst.AddDays(-1)
		$startDate=$firstDayOfMonth
		$endDate=$lastDayoOfMonth
	} else {
		if ($args[0].Contains(".")) {
	        $inputSplit=$args[0].split(".")
			if ($inputSplit.count -gt 1) {
				$inputDate=$inputSplit[1]+"/"+$inputSplit[0]+"/"+$inputSplit[2]			
			} else {
			    Write-Host ".: ERROR: Invalid date input "$args[0]"!"
				Exit 1
			}
		} elseif ($args[0].Contains("/")) {
	        $inputSplit=$args[0].split("/")
			if ($inputSplit.count -gt 1) {
				$inputDate=$inputSplit[0]+"/"+$inputSplit[1]+"/"+$inputSplit[2]			
			} else {
			    Write-Host ".: ERROR: Invalid date input "$args[0]"!"
				Exit 1
			}
		} else {
			try {
		        $result=Get-Date $inputDate
			} catch {
			    Write-Host ".: ERROR: Invalid date input "$args[0]": "$_
				Exit 1
			}
		}
		
		$result=Get-Date $inputDate
   		$startDate=$result
    	$endDate=$result
		
		if ($args.count -gt 1) {
			try {
        		$inputSplit=$args[1].split(".")
	        	if ($inputSplit.count -gt 1) {
					$inputDate=$inputSplit[1]+"/"+$inputSplit[0]+"/"+$inputSplit[2]
				}
				$result=Get-Date $inputDate
				$endDate=$result
			} catch {
			}
		}
	}
}
$month = $startDate.ToString("MM")
$year = $startDate.ToString("yyyy")
$startDateStr = $startDate.ToString($msFormatStr)
$endDateStr = $endDate.ToString($msFormatStr)
$currentPath = Get-Location
$dataPath = "$currentPath\$dataPathName"
$conifgPath = "$currentPath\$configPathName"
$csvPath = "$dataPath\$year$month-mail-export.csv"


###############################################################################
# Main Processing
###############################################################################
Write-Host $LINE
Write-Host "   Export outlook mails from" $startDate.ToString($userFormatStr) "-" $endDate.ToString($userFormatStr)
Write-Host $LINE

try {
	# Get outlook objects
	Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
	$outlook = New-Object -ComObject Outlook.Application
	$namespace = $outlook.GetNamespace("MAPI")

	$currentUserInfo = Get-OutlookCurrentUser -namespace $namespace
	$currentUserDisplayName = $currentUserInfo.DisplayName
	$currentUserEmail = $currentUserInfo.EmailAddress
    Write-Host ".: Outlook User Information: $currentUserDisplayName, mail: $currentUserEmail"

	# Define folders to search
	$folders = @(
	#	@{ Folder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox); Name = "Inbox" },
		@{ Folder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail); Name = "Sent" }
	)

	# Define date filter for emails (using SentOn for consistency)
	$filter = "[SentOn] >= '" + $startDateStr + " 00:00 AM' AND [SentOn] <= '" + $endDateStr + " 23:59 PM'"
	#Write-Host ".: Using email filter: $filter"
	Write-Host ".: Select data..."
	
	# Collect emails (initial collection without duration calculation)
	$emails = @()
	$totalEmailsProcessed = 0

	foreach ($folderInfo in $folders) {
		$folder = $folderInfo.Folder
		$folderName = $folderInfo.Name
		
		Write-Host ".: Processing folder: $folderName"
		
		try {
			$items = $folder.Items.Restrict($filter)
			$items.Sort("[SentOn]", $true)
			
			Write-Host ".: Found $($items.Count) emails in $folderName"
			
			foreach ($item in $items) {
				if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
					$totalEmailsProcessed++
					
					# Get email date and time
					$emailDate = Get-Date "$($item.SentOn)"
					$emailTime = $emailDate.ToString("HH:mm")
					
					# Clean the subject and store both original and cleaned versions
					$originalSubject = $item.Subject
					$cleanedSubject = Clean-EmailSubject -Subject $originalSubject
					
					$emails += [PSCustomObject]@{
						Date          = $emailDate.ToString($userFormatStr)
						Duration      = $DEFAULT_EMAIL_DURATION  # Will be recalculated
						Subject       = $cleanedSubject  # Use cleaned subject for export
						Start         = $emailTime
						End           = ""  # Emails don't have end times
						Location      = $folderName  # Use folder name as "location"
						Attendees     = Get-FullEmailRecipients -item $item
						Private       = if ($item.Sensitivity -eq [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate) { "TRUE" } else { "FALSE" }
						CalendarOwner = $currentUserInfo.DisplayName
						OwnerEmail    = $currentUserInfo.EmailAddress
						OriginalSubject = $originalSubject  # Keep original for reference
						CleanedSubject = $cleanedSubject   # For duration calculation
					}
				}
			}
		} catch {
			Write-Host ".: ERROR processing folder $folderName : $_"
		}
	}

	# Calculate smart durations based on daily subject frequency
	if ($emails.Count -gt 0) {
		$emails = Calculate-EmailDuration -AllEmails $emails

		
		# Group emails by date and subject to create single entries per day
		#Write-Host ".: Consolidating emails by date and subject..."
		$consolidatedEmails = @{}
		$consolidatedCount = 0
		
		foreach ($email in $emails) {
			$groupKey = "$($email.Date)|$($email.CleanedSubject.ToLower())"
			
			if (-not $consolidatedEmails.ContainsKey($groupKey)) {
				# First email for this date/subject combination
				$consolidatedEmails[$groupKey] = [PSCustomObject]@{
					Subject       = $email.CleanedSubject
					Date          = $email.Date
					Start         = $email.Start  # Time of first email
					End           = ""
					Duration      = $email.Duration
					Location      = $email.Location
					Attendees     = $email.Attendees
					Private       = $email.Private
					CalendarOwner = $email.CalendarOwner
					OwnerEmail    = $email.OwnerEmail
					EmailCount    = 1
					TimeRange     = $email.Start
				}
			} else {
				# Additional email for same date/subject - accumulate duration
				$existingEmail = $consolidatedEmails[$groupKey]
				$existingEmail.Duration += $email.Duration
				$existingEmail.EmailCount++
				
				# Update time range to show first-last email time
				if ($email.Start -ne $existingEmail.Start) {
					$existingEmail.TimeRange = "$($existingEmail.Start)-$($email.Start)"
				}
				
				# Merge attendees if different
				if ($email.Attendees -ne $existingEmail.Attendees) {
					$allAttendees = @($existingEmail.Attendees, $email.Attendees) | Where-Object { $_ -and $_.Trim() -ne "" } | Sort-Object -Unique
					$existingEmail.Attendees = $allAttendees -join "; "
				}
				
				$consolidatedCount++
			}
		}
		
		# Convert hashtable to array
		$finalEmails = @()
		foreach ($key in $consolidatedEmails.Keys) {
			$email = $consolidatedEmails[$key]
			
			# Update subject to show email count if more than 1
			if ($email.EmailCount -gt 1) {
				$email.Subject = "$($email.Subject) ($($email.EmailCount) emails)"
			}
			
			# Use time range as Start if we have multiple emails
			if ($email.TimeRange -ne $email.Start) {
				$email.Start = $email.TimeRange
			}
			
			$finalEmails += $email
		}
		
		if ($consolidatedCount -gt 0) {
			#Write-Host ".: Consolidated $consolidatedCount duplicate emails into single entries"
		}
		
		# Replace emails array with consolidated version
		$emails = $finalEmails		
	}


	###############################################################################
	# Export to CSV
	###############################################################################
	if (-not (Test-Path -Path $dataPath)) {
		New-Item -Path $dataPath -ItemType Directory | Out-Null
	}

    # Export to CSV (remove helper properties before export)
	if ($emails.Count -gt 0) {
		$exportEmails = @()
		foreach ($email in $emails) {
			$exportEmails += [PSCustomObject]@{
				Date          = $email.Date
				Duration      = $email.Duration
				Subject       = $email.Subject
				Start         = $email.Start
				End           = $email.End
				Location      = $email.Location
				Attendees     = $email.Attendees
				Private       = $email.Private
				CalendarOwner = $email.CalendarOwner
				OwnerEmail    = $email.OwnerEmail
			}
		}
		
		$exportEmails | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
		Write-Host ".: Email export complete. File saved to: $csvPath"
		Write-Host ".: Exported $($emails.Count) emails from $totalEmailsProcessed total processed"
		
		# Show duration summary
		$totalHours = ($emails | Measure-Object -Property Duration -Sum).Sum
		#Write-Host ".: Total estimated time: $totalHours hours"
	} else {
		Write-Host ".: No emails found in the specified date range"
	}

	# Show cache statistics
	#Show-CacheStats
} catch {
    Write-Host ".: ERROR: $($_.Exception.Message)"
    Exit 1
} finally {
    # Clean up COM objects
    if ($folders) {
        foreach ($folderInfo in $folders) {
            if ($folderInfo.Folder) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folderInfo.Folder) | Out-Null
            }
        }
    }
    if ($namespace) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
    }
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}