###############################################################################
#
# calendar-exporter.ps1
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
$caseSensitive = $true  # Set to $false for case-insensitive filtering
#$msFormatStr = "MM/dd/yyyy"
$msFormatStr = "dd/MM/yyyy"
$userFormatStr = "dd/MM/yyyy"
$dataPathName = "data"
$configPathName = "config"


###############################################################################
# Filter words
###############################################################################
function Get-FilterWords {
    param (
        [string]$RelativePath
    )

    if (Test-Path -Path $RelativePath) {
        $result = Get-Content -Path $RelativePath
        Write-Host "   - Loaded filter file:" $RelativePath
        return $result
    } else {
        #Write-Host ".: No config file found ($RelativePath)."
	    return @()  # Return empty array instead of $null
    }											 
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
function Get-FullAttendees {
    param($item)
    
    $attendeesList = @()
    try {
        if ($item.Recipients.Count -gt 0) {
            for ($i = 1; $i -le $item.Recipients.Count; $i++) {
                $recipient = $item.Recipients.Item($i)
                $cacheKey = "" # Create cache key from recipient properties
                try {
                    $cacheKey = $recipient.AddressEntry.ID # Use EntryID as primary cache key (most reliable)
                } catch {
                    $cacheKey = "$($recipient.Name)|$($recipient.Address)" # Fallback: use name + address combination
                }
                
                # Check cache first
                if ($global:attendeeCache.ContainsKey($cacheKey)) {
                    $attendeesList += $global:attendeeCache[$cacheKey]
                    $global:cacheStats.Hits++
                    continue
                }
                
                $global:cacheStats.Misses++ # Not in cache - resolve attendee info
                $resolvedAttendee = Resolve-SingleAttendee -recipient $recipient
                
                $global:attendeeCache[$cacheKey] = $resolvedAttendee # Cache the result
                $attendeesList += $resolvedAttendee
            }
        }
    } catch {
        $global:cacheStats.Errors++ # Fallback to original method if Recipients collection fails
        return $item.RequiredAttendees 
    }
    return ($attendeesList -join "; ") # Join all attendees with semicolon
}

# Resolve single attendee (separated for caching efficiency)
function Resolve-SingleAttendee {
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
        Write-Host ".: Attendee Cache Stats:"
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
# Duration conversion function - Add this after your Get-FilterWords function
###############################################################################
function Convert-MinutesToRoundedHours {
    param (
        [double]$Minutes
    )
    
    # Convert minutes to hours and round to nearest 0.5
    # Logic: 
    # 0-30 minutes = 0.5 hours
    # 31-60 minutes = 1.0 hours  
    # 61-90 minutes = 1.5 hours
    # 91-120 minutes = 2.0 hours
    # etc.
    
    if ($Minutes -le 0) {
        return 0.0
    } elseif ($Minutes -le 30) {
        return 0.5
    } else {
        # For minutes > 30, round up to nearest 0.5 hour
        $hours = $Minutes / 60.0
        $roundedHours = [Math]::Ceiling($hours * 2) / 2
        return $roundedHours
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
$csvPath = "$dataPath\$year$month-calendar-export.csv"


###############################################################################
# Main
###############################################################################
Write-Host $LINE
Write-Host "   Export outlook calendar from" $startDate.ToString($userFormatStr) "-" $endDate.ToString($userFormatStr)
Write-Host $LINE

$outlook = $null
$namespace = $null
$calendar = $null

try {
	# Get outlook objects
	Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
	$outlook = New-Object -ComObject Outlook.Application
	$namespace = $outlook.GetNamespace("MAPI")

	$currentUserInfo = Get-OutlookCurrentUser -namespace $namespace
	$currentUserDisplayName = $currentUserInfo.DisplayName
	$currentUserEmail = $currentUserInfo.EmailAddress
    Write-Host ".: Outlook User Information: $currentUserDisplayName, mail: $currentUserEmail"
	#Show-CurrentUserInfo -userInfo $currentUserInfo

	Write-Host ".: Check config files..."
	$calendarSubjectFilter = Get-FilterWords -RelativePath "$conifgPath\calendar-subject-filter.txt"
	if (-not $caseSensitive -and $calendarSubjectFilter.Count -gt 0) {
		$calendarSubjectFilter = $calendarSubjectFilter | ForEach-Object { $_.ToLower() }
	}

	$calendarAttendeeFilter = Get-FilterWords -RelativePath "$conifgPath\calendar-attendee-filter.txt"
	if (-not $caseSensitive -and $calendarAttendeeFilter.Count -gt 0) {
		$calendarAttendeeFilter = $calendarAttendeeFilter | ForEach-Object { $_.ToLower() }
	}

	# Get Calendar folder
	$calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
	$calendarItems = $calendar.Items
	$calendarItems.Sort("[Start]", $true)
	$calendarItems.IncludeRecurrences = $true

	# Format dates for Outlook restriction
	$calendarEntryFilter = "[Start] >= '"+$startDateStr+" 00:00 AM' AND [Start] <= '"+$endDateStr+" 23:59 PM'"
	#$calendarEntryFilter = "[Start] >= '$startDateStr' AND [Start] <= '$endDateStr'"
	#Write-Host ".: Using email filter: $filter"
	Write-Host ".: Select data..."
	$items = $calendarItems.Restrict($calendarEntryFilter)

	# Collect appointments
	$appointments = @()
	foreach ($item in $items) {
		if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem] -and -not $item.AllDayEvent) {
			# Skip if subject matches any filter word (case-insensitive)
			$subjectToCheck = if ($caseSensitive) { $item.Subject } else { $item.Subject.ToLower() }
			if ($calendarSubjectFilter -contains $subjectToCheck) {
				continue
			}
			# Skip if subject matches any filter word (case-insensitive)
			$attendeesToCheck = if ($caseSensitive) { $item.RequiredAttendees } else { $item.RequiredAttendees.ToLower() }
			if ($calendarAttendeeFilter -contains $attendeesToCheck) {
				continue
			}
			
			$itemDate = Get-Date "$($item.Start)"
		#	if (($firstDayOfMonth -ge $itemDate) -or ($itemDate -ge $lastDayoOfMonth)) {
				#Write-Host "Skip" $itemDate.ToString($userFormatStr)
		#	} else {
			
			$itemStartTimestamp = (Get-Date "$($item.Start)").ToString("HH:mm")
			$itemEndTimestamp = (Get-Date "$($item.End)").ToString("HH:mm")
		    #Write-Host ".: Outlook User Information: $($currentUserInfo.DisplayName), mail: $($currentUserInfo.EmailAddress)"

			if ($item.IsRecurring) {
				$pattern = $item.GetRecurrencePattern()
				$timeStr = $item.Start.ToString("HH:mm")+":00"
				
				$currentDate = $startDate
				while ($currentDate -le $endDate) {
					try {
						
						$checkDate = $currentDate.ToString("MM/dd/yyyy")+" $timeStr"
						$occurenceDate = Get-Date "$checkDate"
						$occurrence = $pattern.GetOccurrence($occurenceDate)
						#Write-Output "$($occurrence.Start.ToString($userFormatStr+" HH:mm")) - $($occurrence.Subject) (recurring)"
						$appointments += [PSCustomObject]@{
							Date      = $occurenceDate.ToString($userFormatStr)
							Duration  = Convert-MinutesToRoundedHours -Minutes ($item.End - $item.Start).TotalMinutes
							#Duration  = ($item.End - $item.Start).TotalMinutes  # or use .TotalHours
							Subject   = $item.Subject
							Start     = $itemStartTimestamp
							End       = $itemEndTimestamp
							Location  = $item.Location
							#Attendees = $item.RequiredAttendees
							Attendees = Get-FullAttendees -item $item
							Private   = $item.Sensitivity -eq [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate
							Recurring = "TRUE"
							CalendarOwner = $currentUserDisplayName
							OwnerEmail    = $currentUserEmail
						}					
					} catch {
						#Write-Output "$checkDate - No occurrence on that date for: $($item.Subject)"
					}
					#try {
					$currentDate = $currentDate.AddDays(1)
					#} catch {
					#}
				}
			} else {
				#Write-Output "$($item.Start.ToString($userFormatStr+" HH:mm")) - $($item.Subject)"
				$appointments += [PSCustomObject]@{
					Date      = $itemDate.ToString($userFormatStr)
					#Duration  = ($item.End - $item.Start).TotalMinutes  # or use .TotalHours
					Duration  = Convert-MinutesToRoundedHours -Minutes ($item.End - $item.Start).TotalMinutes
					Subject   = $item.Subject
					Start     = $itemStartTimestamp
					End       = $itemEndTimestamp
					Location  = $item.Location
					#Attendees = $item.RequiredAttendees
					Attendees = Get-FullAttendees -item $item
					Private   = $item.Sensitivity -eq [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate
					Recurring   = "FALSE"
					CalendarOwner = $currentUserDisplayName
					OwnerEmail    = $currentUserEmail
				}
			}
		}
	}


	###############################################################################
	# Export to CSV
	###############################################################################
	if (-not (Test-Path -Path $dataPath)) {
		New-Item -Path $dataPath -ItemType Directory | Out-Null
	}
	$appointments | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

	Write-Host ".: Calendar export complete. File saved to: $csvPath"
	#Show-CacheStats
	Write-Host ".: Exported $($appointments.Count) appointments"
} catch {
    Write-Host ".: ERROR: $($_.Exception.Message)"
    Exit 1
} finally {
    # Clean up COM objects
    if ($calendar) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
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