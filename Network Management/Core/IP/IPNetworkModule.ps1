# ===================================================================
# IP NETWORK IDENTIFIER MODULE - COMPLETE VERSION
# ===================================================================

# Search debouncing timer
$script:IPSearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:IPSearchTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$script:IPSearchTimer.Add_Tick({
    Update-SubnetDataGridWithSearch
    $script:IPSearchTimer.Stop()
})

class SubnetEntry {
    [int]$ID
    [string]$IP_Subnet
    [int]$VLAN_ID
    [string]$VLAN_Name
    [string]$Site_Name

    SubnetEntry([string]$subnet, [int]$vlanId, [string]$vlanName, [string]$siteName) {
        if ([string]::IsNullOrWhiteSpace($subnet) -or $subnet -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}$') {
            throw "Invalid subnet format: $subnet"
        }
        $this.IP_Subnet = $subnet
        $this.VLAN_ID = $vlanId
        $this.VLAN_Name = if($vlanName) { $vlanName } else { "" }
        $this.Site_Name = if($siteName) { $siteName } else { "" }
    }
}

class SubnetDataStore {
    hidden [string]$DataFile = "$(Split-Path $PSScriptRoot -Parent)\Data\ip_data.json"
    hidden [System.Collections.Generic.List[SubnetEntry]]$Entries

    SubnetDataStore() {
        $this.Entries = [System.Collections.Generic.List[SubnetEntry]]::new()
        $this.LoadData()
    }

    [void] LoadData() {
        try {
            if (Test-Path $this.DataFile) {
                $jsonData = Get-Content $this.DataFile -Raw | ConvertFrom-Json
                $this.Entries.Clear()
                
                if ($jsonData) {
                    foreach ($item in $jsonData) {
                        if ($item -and $item.IP_Subnet) {
                            $entry = [SubnetEntry]::new(
                                $item.IP_Subnet,
                                $item.VLAN_ID,
                                $item.VLAN_Name,
                                $item.Site_Name
                            )
                            $entry.ID = $item.ID
                            $this.Entries.Add($entry)
                        }
                    }
                }
            }
        } catch {
            $this.Entries.Clear()
        }
    }

    [void] SaveData() {
        try {
            if ($this.Entries.Count -eq 0) {
                if (Test-Path $this.DataFile) {
                    Remove-Item $this.DataFile -Force
                }
            } else {
                $this.Entries | ConvertTo-Json -Depth 3 | Set-Content $this.DataFile -Encoding UTF8
            }
        } catch {
        }
    }

    [SubnetEntry[]] GetAllEntries() {
        if ($this.Entries -eq $null -or $this.Entries.Count -eq 0) {
            return @()
        }
        return $this.Entries.ToArray()
    }

    [bool] AddEntry([SubnetEntry]$entry) {
        try {
            if ($entry -eq $null) { return $false }
            
            # Check for duplicates
            foreach ($existing in $this.Entries) {
                if ($existing.IP_Subnet -eq $entry.IP_Subnet) {
                    return $false
                }
            }
            
            $entry.ID = $this.GetNextAvailableId()
            $this.Entries.Add($entry)
            $this.SaveData()
            return $true
        } catch {
            return $false
        }
    }

    [bool] DeleteEntries([int[]]$ids) {
        try {
            if ($ids -eq $null -or $ids.Count -eq 0) { return $false }
            
            $countBefore = $this.Entries.Count
            $newList = [System.Collections.Generic.List[SubnetEntry]]::new()
            
            foreach ($entry in $this.Entries) {
                if ($entry.ID -notin $ids) {
                    $newList.Add($entry)
                }
            }
            
            $this.Entries = $newList
            $this.SaveData()
            return $this.Entries.Count -lt $countBefore
        } catch {
            return $false
        }
    }

    hidden [int] GetNextAvailableId() {
        if ($this.Entries.Count -eq 0) { return 1 }
        $maxId = 0
        foreach ($entry in $this.Entries) {
            if ($entry.ID -gt $maxId) { $maxId = $entry.ID }
        }
        return $maxId + 1
    }
}

# Validation functions to reduce code duplication
function Test-IPSubnetFormat {
    param([string]$IPSubnet)
    
    try {
        # Check if input is null or empty
        if ([string]::IsNullOrWhiteSpace($IPSubnet)) {
            return $false
        }
        
        # Basic format check
        if ($IPSubnet -notmatch '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})/(\d{1,2})$') {
            return $false
        }
        
        # Split IP and CIDR parts
        $parts = $IPSubnet.Split('/')
        if ($parts.Count -ne 2) {
            return $false
        }
        
        $ipPart = $parts[0]
        $cidrPart = $null
        
        # Safely parse CIDR
        if (-not [int]::TryParse($parts[1], [ref]$cidrPart)) {
            return $false
        }
        
        # Validate CIDR range (0-32 for IPv4)
        if ($cidrPart -lt 0 -or $cidrPart -gt 32) {
            return $false
        }
        
        # Validate each IP octet (0-255)
        $octets = $ipPart.Split('.')
        if ($octets.Count -ne 4) {
            return $false
        }
        
        foreach ($octet in $octets) {
            $octetValue = $null
            if (-not [int]::TryParse($octet, [ref]$octetValue)) {
                return $false
            }
            if ($octetValue -lt 0 -or $octetValue -gt 255) {
                return $false
            }
        }
        
        # Additional validation: Parse as IPAddress to catch edge cases
        $ipAddress = $null
        if (-not [System.Net.IPAddress]::TryParse($ipPart, [ref]$ipAddress)) {
            return $false
        }
        
        # Validate that it's a proper network address (not a host address)
        # Calculate network address and compare with input
        $ip = [System.Net.IPAddress]::Parse($ipPart)
        $mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $cidrPart))
        $maskBytes = [BitConverter]::GetBytes($mask)
        $ipBytes = $ip.GetAddressBytes()
        
        # Calculate network address
        $networkBytes = @()
        for ($i = 0; $i -lt 4; $i++) {
            $networkBytes += ($ipBytes[$i] -band $maskBytes[$i])
        }
        
        # Check if provided IP matches the network address
        for ($i = 0; $i -lt 4; $i++) {
            if ($ipBytes[$i] -ne $networkBytes[$i]) {
                # This is a host address, not a network address
                return $false
            }
        }
        
        return $true
    } catch {
        return $false
    }
}

function Test-VlanId {
    param([string]$VlanId)
    
    try {
        if ([string]::IsNullOrWhiteSpace($VlanId)) {
            return $false
        }
        
        $result = $null
        if ([int]::TryParse($VlanId.Trim(), [ref]$result)) {
            return $result -ge 1 -and $result -le 4094  # Valid VLAN range
        }
        return $false
        
    } catch {
        return $false
    }
}

function Show-IPValidationError {
    param(
        [string]$Message,
        [string]$Title = "Validation Error"
    )
    
    # Use the import status control since txtBlkStatus might not exist
    if ($txtBlkImportStatus -ne $null) {
        try {
            $txtBlkImportStatus.Text = $Message
            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
        } catch {
        }
    }
    
    # Show dialog
    Show-CustomDialog $Message $Title "OK" "Error"
}

# Column name constants to avoid hardcoded strings
$COLUMN_IP_SUBNET = "IP Subnet"
$COLUMN_VLAN_ID = "VLAN ID" 
$COLUMN_VLAN_NAME = "VLAN Name"
$COLUMN_SITE_NAME = "Site Name"

# Security function to prevent CSV injection
function Remove-CSVInjection {
    param([string]$Value)
    
    try {
        if ([string]::IsNullOrWhiteSpace($Value)) { 
            return "" 
        }
        
        # Remove dangerous characters that could cause CSV injection
        $dangerous = @('=', '+', '-', '@', '|', '%')
        foreach ($char in $dangerous) {
            if ($Value.StartsWith($char)) {
                $Value = "'" + $Value  # Prefix with quote to neutralize
            }
        }
        
        # Remove control characters and limit length
        $Value = $Value -replace '[\x00-\x1F\x7F]', ''
        return $Value.Substring(0, [Math]::Min(255, $Value.Length))
    } catch {
        return ""
    }
}

function Add-SubnetEntry {
    param (
        [string]$IpSubnet,
        [int]$VlanId,
        [string]$VlanName,
        [string]$SiteName
    )
    
    try {
        # Validate inputs are not null
        if ([string]::IsNullOrWhiteSpace($IpSubnet)) {
            throw "IP Subnet cannot be empty"
        }
        if ([string]::IsNullOrWhiteSpace($VlanName)) {
            throw "VLAN Name cannot be empty"
        }
        if ([string]::IsNullOrWhiteSpace($SiteName)) {
            throw "Site Name cannot be empty"
        }
        if ($VlanId -le 0) {
            throw "VLAN ID must be a positive number"
        }
        
        # Check if subnetDataStore exists
        if ($subnetDataStore -eq $null) {
            throw "Subnet data store is not initialized"
        }
        
        $entry = [SubnetEntry]::new($IpSubnet, $VlanId, $VlanName, $SiteName)
        
        if ($subnetDataStore.AddEntry($entry)) {
            $successMsg = "Entry added successfully! (ID: $($entry.ID))"
            
            # Safely update status if control exists
            if ($txtBlkImportStatus -ne $null) {
                try {
                    $txtBlkImportStatus.Text = $successMsg
                    $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                } catch {
                }
            }
            return $true
        } else {
            Show-IPValidationError "Error: Subnet already exists"
            return $false
        }
    } catch {
        Show-IPValidationError "Error: $($_.Exception.Message)"
        return $false
    }
}

function Lookup-IpAddress {
    param (
        [string]$IpAddress
    )

    try {
        # Always hide results at start
        if ($grpLookupResults -ne $null) {
            $grpLookupResults.Visibility = "Collapsed"
        }
        
        # Clear previous results
        if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = "" }
        if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = "" }
        if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = "" }
        if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = "" }
        if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = "" }

        if ([string]::IsNullOrEmpty($IpAddress)) {
            Show-CustomDialog "Please enter an IP address" "Input Required" "OK" "Warning"
            return
        }

        try {
            $ip = [System.Net.IPAddress]::Parse($IpAddress)
        } catch {
            Show-CustomDialog "Invalid IP address format" "Input Error" "OK" "Error"
            return
        }

        $foundMatch = $false
        $allEntries = $subnetDataStore.GetAllEntries()
        
        foreach ($entry in $allEntries) {
            if ($entry -eq $null) { continue }
            
            $subnetCidr = $entry.IP_Subnet
            if (-not $subnetCidr -or $subnetCidr -notmatch '^(\d{1,3}\.){3}\d{1,3}/\d{1,2}$') {
                continue
            }

            $parts = $subnetCidr.Split('/')
            try {
                $subnetIp = [System.Net.IPAddress]::Parse($parts[0])
                $prefixLength = [int]$parts[1]
                
                $mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $prefixLength))
                $maskBytes = [BitConverter]::GetBytes($mask)
                
                $ipBytes = $ip.GetAddressBytes()
                $subnetIpBytes = $subnetIp.GetAddressBytes()
                
                $isInSubnet = $true
                for ($i = 0; $i -lt 4; $i++) {
                    if (($ipBytes[$i] -band $maskBytes[$i]) -ne ($subnetIpBytes[$i] -band $maskBytes[$i])) {
                        $isInSubnet = $false
                        break
                    }
                }

                if ($isInSubnet) {
                    # Only show results on successful match
                    if ($grpLookupResults -ne $null) { $grpLookupResults.Visibility = "Visible" }
                    if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = $IpAddress }
                    if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = $subnetCidr }
                    if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = $entry.VLAN_ID }
                    if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = $entry.VLAN_Name }
                    if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = $entry.Site_Name }
                    $foundMatch = $true
                    break
                }
            } catch {
                continue
            }
        }

        if (-not $foundMatch) {
            Show-CustomDialog "IP address $IpAddress not found in the network database." "Not Found" "OK" "Information"
        }
    } catch {
        Show-CustomDialog "Error during IP lookup: $_" "Error" "OK" "Error"
    }
}

function Import-SubnetsFromCsv {
    param (
        [string]$CsvFilePath
    )

    try {
        # Basic file path validation
        if ([string]::IsNullOrWhiteSpace($CsvFilePath)) {
            throw "File path cannot be empty"
        }
        
        if (-not (Test-Path $CsvFilePath)) {
            throw "CSV file not found"
        }
        
        # Check file extension
        if (-not $CsvFilePath.ToLower().EndsWith('.csv')) {
            throw "File must be a CSV file (.csv extension)"
        }
        
        # Check file size (prevent loading huge files)
        $fileInfo = Get-Item $CsvFilePath
        if ($fileInfo.Length -gt 50MB) {
            throw "File is too large. Maximum size is 50MB"
        }

        # Show progress panel
        if ($pnlImportProgress -ne $null) {
            $pnlImportProgress.Visibility = [System.Windows.Visibility]::Visible
        }
        if ($pbImportProgress -ne $null) { $pbImportProgress.Value = 0 }
        if ($txtProgressStatus -ne $null) { $txtProgressStatus.Text = "Starting import..." }
        if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "" }

        # Read CSV data
        $csvData = Import-Csv -Path $CsvFilePath
        $totalLines = $csvData.Count
        $importedCount = 0
        $skippedCount = 0
        $errorMessages = @()

        foreach ($row in $csvData) {
            try {
                # Validate required fields
                if ([string]::IsNullOrWhiteSpace($row.IP_Subnet) -or 
                    [string]::IsNullOrWhiteSpace($row.VLAN_ID) -or 
                    [string]::IsNullOrWhiteSpace($row.VLAN_Name) -or 
                    [string]::IsNullOrWhiteSpace($row.Site_Name)) {
                    throw "Missing required fields"
                }

                # Validate IP format
                if (-not (Test-IPSubnetFormat $row.IP_Subnet)) {
                    throw "Invalid subnet format: $($row.IP_Subnet)"
                }

                # Check for duplicates
                $existingEntries = $subnetDataStore.GetAllEntries()
                if ($existingEntries.IP_Subnet -contains $row.IP_Subnet) {
                    throw "Duplicate subnet: $($row.IP_Subnet)"
                }

                # Sanitize data to prevent CSV injection attacks
                $cleanIPSubnet = Remove-CSVInjection $row.IP_Subnet
                $cleanVlanName = Remove-CSVInjection $row.VLAN_Name
                $cleanSiteName = Remove-CSVInjection $row.Site_Name

                # Create and add entry with sanitized data
                $entry = [SubnetEntry]::new(
                    $cleanIPSubnet,
                    $row.VLAN_ID,
                    $cleanVlanName,
                    $cleanSiteName
                )
                
                if ($subnetDataStore.AddEntry($entry)) {
                    $importedCount++
                }

                # Update progress
                $totalProcessed = $importedCount + $skippedCount
                if ($totalProcessed % 10 -eq 0 -or $totalProcessed -eq $totalLines) {
                    $progress = [math]::Round(($totalProcessed / $totalLines) * 100)
                    if ($pbImportProgress -ne $null) { $pbImportProgress.Value = $progress }
                    if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "$importedCount / $totalLines records processed" }
                    [System.Windows.Forms.Application]::DoEvents()
                }

            } catch {
                $skippedCount++
                $errorMessages += "Line $($importedCount + $skippedCount): $_"
            }
        }

        $result = @(
            "Import completed",
            "Successfully imported: $importedCount",
            "Skipped: $skippedCount"
        ) -join "`n"

        if ($errorMessages.Count -gt 0) {
            $result += "`n`nErrors:`n" + ($errorMessages -join "`n")
        }

        return $result
    } catch {
        throw "Import failed: $_"
    } finally {
        if ($pnlImportProgress -ne $null) {
            $pnlImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }
}

function Export-SubnetsToCsv {
    param (
        [string]$FilePath
    )
    
    try {
        # Basic file path validation
        if ([string]::IsNullOrWhiteSpace($FilePath)) {
            throw "File path cannot be empty"
        }
        
        $data = $subnetDataStore.GetAllEntries() | Select-Object IP_Subnet, VLAN_ID, VLAN_Name, Site_Name
        $data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        return "Exported $($data.Count) subnet entries to:`n$FilePath"
    } catch {
        throw "Export failed: $($_.Exception.Message)"
    }
}

function Update-SubnetDataGridWithSearch {
    try {
        $searchTerm = if ($txtSearch -ne $null) { $txtSearch.Text } else { "" }
        
        # Get all data from the data store
        $allData = @($subnetDataStore.GetAllEntries())
        
        # Filter data if search term exists
        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            $searchTerm = $searchTerm.Trim().ToLower()
            $filteredData = $allData | Where-Object {
                $_.IP_Subnet.ToLower().Contains($searchTerm) -or
                $_.VLAN_ID.ToString().Contains($searchTerm) -or
                $_.VLAN_Name.ToLower().Contains($searchTerm) -or
                $_.Site_Name.ToLower().Contains($searchTerm)
            }
            # Force array conversion to ensure we always have an IEnumerable
            $allData = @($filteredData)
        }
        
        # Sort by ID numerically and update DataGrid
        if ($allData.Count -gt 0) {
            $allData = @($allData | Sort-Object -Property @{Expression={[int]$_.ID}; Ascending=$true})
        }
        
        # Safely update DataGrid - ensure we always pass an array/collection
        if ($dgSubnets -ne $null) {
            # Convert to ArrayList to ensure proper IEnumerable interface
            $dataGridSource = New-Object System.Collections.ArrayList
            foreach ($item in $allData) {
                $null = $dataGridSource.Add($item)
            }
            $dgSubnets.ItemsSource = $dataGridSource
        }
        
        # Update status bar only if values changed
        $newCount = $allData.Count
        if ($txtStatusBarSubnets -ne $null) {
            $txtStatusBarSubnets.Text = "Total Subnets: $newCount"
        }
        if ($txtStatusBarSelected -ne $null) {
            $txtStatusBarSelected.Text = "Selected: None"
        }
        
    } catch {
        
        # Ensure DataGrid is not null even on error
        if ($dgSubnets -ne $null) {
            $dgSubnets.ItemsSource = New-Object System.Collections.ArrayList
        }
    }
}

function Initialize-IPNetworkEventHandlers {
    try {
        # --- Event Handlers ---
        if ($btnAddEntry -ne $null) {
            $btnAddEntry.Add_Click({
                try {
                    # Safely get values with null checks
                    $ipSubnet = ""
                    $vlanId = ""
                    $vlanName = ""
                    $siteName = ""
                    
                    if ($txtIpSubnet -ne $null) { $ipSubnet = $txtIpSubnet.Text.Trim() }
                    if ($txtVlanId -ne $null) { $vlanId = $txtVlanId.Text.Trim() }
                    if ($txtVlanName -ne $null) { $vlanName = $txtVlanName.Text.Trim() }
                    if ($txtSiteName -ne $null) { $siteName = $txtSiteName.Text.Trim() }

                    # Validate all fields are filled
                    if ([string]::IsNullOrEmpty($ipSubnet) -or 
                        [string]::IsNullOrEmpty($vlanId) -or 
                        [string]::IsNullOrEmpty($vlanName) -or 
                        [string]::IsNullOrEmpty($siteName)) {
                        
                        $errorMsg = "Error: All fields must be filled."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Validate VLAN ID is numeric
                    if (-not (Test-VlanId $vlanId)) {
                        $errorMsg = "Error: VLAN ID must be a valid number between 1 and 4094."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Validate IP Subnet format (CIDR notation)
                    if (-not (Test-IPSubnetFormat $ipSubnet)) {
                        $errorMsg = "Error: Invalid IP Subnet format. Use CIDR notation (e.g., 10.10.10.0/24)."
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $errorMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                        }
                        Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                        return
                    }

                    # Check for duplicate subnets
                    if ($subnetDataStore -ne $null) {
                        $currentData = @($subnetDataStore.GetAllEntries())
                        if ($currentData | Where-Object { $_.IP_Subnet -eq $ipSubnet }) {
                            $errorMsg = "Error: The subnet '$ipSubnet' already exists in the database."
                            if ($txtBlkImportStatus -ne $null) {
                                $txtBlkImportStatus.Text = $errorMsg
                                $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                            }
                            Show-CustomDialog $errorMsg "Validation Error" "OK" "Error"
                            return
                        }
                    }

                    # Try to add the entry
                    if (Add-SubnetEntry -IpSubnet $ipSubnet -VlanId ([int]$vlanId) -VlanName $vlanName -SiteName $siteName) {
                        $successMsg = "New subnet added successfully!"
                        
                        # Use safe status update
                        if ($txtBlkImportStatus -ne $null) {
                            $txtBlkImportStatus.Text = $successMsg
                            $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                        }
                        
                        Show-CustomDialog $successMsg "Success" "OK" "Information"
                        
                        # Reset form safely
                        if ($txtIpSubnet -ne $null) { $txtIpSubnet.Text = "" }
                        if ($txtVlanId -ne $null) { $txtVlanId.Text = "" }
                        if ($txtVlanName -ne $null) { $txtVlanName.Text = "" }
                        if ($txtSiteName -ne $null) { $txtSiteName.Text = "" }
                        
                        Update-SubnetDataGridWithSearch
                    }
                    
                } catch {
                    if ($txtBlkImportStatus -ne $null) {
                        $txtBlkImportStatus.Text = "Error adding entry: $_"
                        $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                    }
                    Show-CustomDialog "Error adding entry: $_" "Error" "OK" "Error"
                }
            })
        }

        if ($btnLookup -ne $null) {
            $btnLookup.Add_Click({
                try {
                    $ipToLookup = if ($txtIpLookup -ne $null) { $txtIpLookup.Text.Trim() } else { "" }
                    Lookup-IpAddress -IpAddress $ipToLookup
                } catch {
                    Show-CustomDialog "Error during lookup: $_" "Error" "OK" "Error"
                }
            })
        }

        if ($btnDeleteEntry -ne $null) {
            $btnDeleteEntry.Add_Click({
                try {
                    if ($dgSubnets -eq $null) {
                        return
                    }
                    
                    $selectedItems = @($dgSubnets.SelectedItems)
                    if ($selectedItems.Count -gt 0) {
                        $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected entries?" "Confirm Deletion" "YesNo" "Warning"
                        
                        if ($confirm -eq "Yes") {
                            # Get IDs safely
                            $idsToDelete = @()
                            foreach ($item in $selectedItems) {
                                if ($item -ne $null -and $item.ID -ne $null) {
                                    $idsToDelete += $item.ID
                                }
                            }
                            
                            if ($idsToDelete.Count -gt 0) {
                                if ($subnetDataStore.DeleteEntries($idsToDelete)) {
                                    # Clear selection before updating
                                    $dgSubnets.SelectedItems.Clear()
                                    
                                    # Update DataGrid
                                    Update-SubnetDataGridWithSearch
                                    
                                    # Use safe status update
                                    $successMsg = "Successfully deleted $($idsToDelete.Count) entries."
                                    if ($txtBlkImportStatus -ne $null) {
                                        $txtBlkImportStatus.Text = $successMsg
                                        $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                                    }
                                    } else {
                                   $errorMsg = "Error deleting entries."
                                   if ($txtBlkImportStatus -ne $null) {
                                       $txtBlkImportStatus.Text = $errorMsg
                                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                   }
                               }
                           }
                       }
                   } else {
                       $warningMsg = "Please select one or more entries to delete."
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = $warningMsg
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                       }
                   }
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Error during delete operation: $_"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
                   
                   # Try to recover by refreshing the DataGrid
                   try {
                       Update-SubnetDataGridWithSearch
                   } catch {
                   }
               }
           })
       }

       if ($dgSubnets -ne $null) {
           $dgSubnets.Add_CellEditEnding({
               param($sender, $e)
               
               try {
                   if ($e.EditAction -ne [System.Windows.Controls.DataGridEditAction]::Commit) {
                       return
                   }

                   $editedItem = $e.Row.Item
                   $column = $e.Column.Header
                   $newValue = ""

                   if ($e.EditingElement -is [System.Windows.Controls.TextBox]) {
                       $newValue = $e.EditingElement.Text
                   }

                   # Get all data from the data store
                   $currentData = $subnetDataStore.GetAllEntries()
                   $originalItem = $currentData | Where-Object { $_.ID -eq $editedItem.ID }

                   if ($column -eq $COLUMN_IP_SUBNET -and -not (Test-IPSubnetFormat $newValue)) {
                       Show-CustomDialog "Invalid IP Subnet format! Use CIDR notation (e.g., 10.10.10.0/24)" "Error" "OK" "Error"
                       $e.EditingElement.Text = $originalItem.IP_Subnet
                       $e.Cancel = $true
                       $dgSubnets.CancelEdit()
                       $dgSubnets.CommitEdit()
                       return
                   }
                   
                   if ($column -eq $COLUMN_VLAN_ID -and -not (Test-VlanId $newValue)) {
                       Show-CustomDialog "VLAN ID must be a number between 1 and 4094!" "Error" "OK" "Error"
                       $originalValue = $originalItem.VLAN_ID
                       $e.EditingElement.Text = $originalValue
                       $e.Cancel = $true
                       $dgSubnets.CancelEdit()
                       $dgSubnets.CommitEdit()
                       return
                   }
                   
                   # Check for duplicates only when editing IP Subnet column
                   if ($column -eq $COLUMN_IP_SUBNET) {
                       $duplicateEntry = $currentData | Where-Object { $_.IP_Subnet -eq $newValue -and $_.ID -ne $editedItem.ID }
                       if ($duplicateEntry) {
                           Show-CustomDialog "Error: IP Subnet '$newValue' already exists in another entry." "Duplicate Subnet" "OK" "Error"
                           $originalValue = $originalItem.IP_Subnet
                           $e.EditingElement.Text = $originalValue
                           $e.Cancel = $true
                           $dgSubnets.CancelEdit()
                           $dgSubnets.CommitEdit()
                           return
                       }
                   }
                   
                   # Find the item to update by ID
                   $itemToUpdate = $null
                   $itemIndex = -1
                   
                   for ($i = 0; $i -lt $currentData.Count; $i++) {
                       if ($currentData[$i].ID -eq $editedItem.ID) {
                           $itemToUpdate = $currentData[$i]
                           $itemIndex = $i
                           break
                       }
                   }
                   
                   if ($itemIndex -ge 0) {
                       # Update the specific property based on column
                       switch ($column) {
                           $COLUMN_IP_SUBNET { $currentData[$itemIndex].IP_Subnet = $newValue }
                           $COLUMN_VLAN_ID { $currentData[$itemIndex].VLAN_ID = [int]$newValue }
                           $COLUMN_VLAN_NAME { $currentData[$itemIndex].VLAN_Name = $newValue }
                           $COLUMN_SITE_NAME { $currentData[$itemIndex].Site_Name = $newValue }
                       }
                       
                       # Save the updated data
                       $subnetDataStore.Entries = [System.Collections.Generic.List[SubnetEntry]]::new()
                       $subnetDataStore.Entries.AddRange($currentData)
                       $subnetDataStore.SaveData()
                       
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Entry updated successfully!"
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                       }
                   } else {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Error: Could not find entry to update."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                       }
                       $e.Cancel = $true
                   }
               } catch {
                   Show-CustomDialog "Error updating entry: $($_.Exception.Message)" "Error" "OK" "Error"
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Error updating entry: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
                   $e.Cancel = $true
               }
           })

           # DataGrid selection changed event
           $dgSubnets.Add_SelectionChanged({
               try {
                   if ($txtStatusBarSelected -ne $null) {
                       $selectedItems = $dgSubnets.SelectedItems
                       if ($selectedItems.Count -gt 0) {
                           if ($selectedItems.Count -eq 1) {
                               $txtStatusBarSelected.Text = "Selected: $($selectedItems[0].IP_Subnet) (ID: $($selectedItems[0].ID))"
                           } else {
                               $txtStatusBarSelected.Text = "Selected: $($selectedItems.Count) entries"
                           }
                       } else {
                           $txtStatusBarSelected.Text = "Selected: None"
                       }
                   }
               } catch {
               }
           })

           # Enable Delete key to remove selected entries
           $dgSubnets.Add_PreviewKeyDown({
               param($sender, $e)
               
               try {
                   # Check if Delete key was pressed
                   if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                       $selectedItems = @($dgSubnets.SelectedItems)
                       
                       if ($selectedItems.Count -gt 0) {
                           # Call the delete function directly (bypassing the button)
                           if ($btnDeleteEntry -ne $null) {
                               $btnDeleteEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                           }
                           
                           # Mark the event as handled to prevent default behavior
                           $e.Handled = $true
                       }
                   }
               } catch {
               }
           })
       }

       if ($btnBrowseCsv -ne $null) {
           $btnBrowseCsv.Add_Click({
               try {
                   # Disable button during operation
                   $btnBrowseCsv.IsEnabled = $false
                   if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "Selecting file..." }
                   
                   # Create and configure file dialog
                   $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                   $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                   $openFileDialog.Title = "Select CSV File to Import"
                   $openFileDialog.CheckFileExists = $false
                   $openFileDialog.CheckPathExists = $false
                   
                   # Show dialog and get result
                   $result = $openFileDialog.ShowDialog()
                   
                   if ($result -eq "OK") {
                       # Just set the path without any validation
                       if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = $openFileDialog.FileName }
                   } else {
                       if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "" }
                   }
               } catch {
                   Show-CustomDialog "Error selecting file: $($_.Exception.Message)" "Error" "OK" "Error"
                   if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text = "" }
               } finally {
                   if ($openFileDialog) {
                       $openFileDialog.Dispose()
                   }
                   $btnBrowseCsv.IsEnabled = $true
               }
           })
       }

       # Optimized Import CSV Button Click Handler
       if ($btnImportCsv -ne $null) {
           $btnImportCsv.Add_Click({
               try {
                   $csvPath = if ($txtCsvFilePath -ne $null) { $txtCsvFilePath.Text.Trim() } else { "" }
                   if ([string]::IsNullOrEmpty($csvPath)) {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Please select a CSV file first."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                       }
                       return
                   }

                   # Disable UI during import
                   $btnImportCsv.IsEnabled = $false
                   if ($btnBrowseCsv -ne $null) { $btnBrowseCsv.IsEnabled = $false }
                   if ($MainTabControl -ne $null) { $MainTabControl.IsEnabled = $false }
                   
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Starting import..."
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                   }

                   # Show progress panel
                   if ($pnlImportProgress -ne $null) {
                       $pnlImportProgress.Visibility = [System.Windows.Visibility]::Visible
                   }
                   if ($pbImportProgress -ne $null) { $pbImportProgress.Value = 0 }
                   if ($txtProgressStatus -ne $null) { $txtProgressStatus.Text = "Preparing import..." }
                   if ($txtProgressDetails -ne $null) { $txtProgressDetails.Text = "" }
                   
                   $result = Import-SubnetsFromCsv -CsvFilePath $csvPath
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = $result
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                   }
                   Update-SubnetDataGridWithSearch
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "Import failed: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
               } finally {
                   # Hide progress panel
                   if ($pnlImportProgress -ne $null) {
                       $pnlImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
                   }
                   
                   # Re-enable UI
                   $btnImportCsv.IsEnabled = $true
                   if ($btnBrowseCsv -ne $null) { $btnBrowseCsv.IsEnabled = $true }
                   if ($MainTabControl -ne $null) { $MainTabControl.IsEnabled = $true }
               }
           })
       }

       if ($btnExportCsv -ne $null) {
           $btnExportCsv.Add_Click({
               try {
                   $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                   $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                   $saveDialog.FileName = "SubnetExport-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').csv"

                   if ($saveDialog.ShowDialog() -eq "OK") {
                       $btnExportCsv.IsEnabled = $false
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = "Exporting..."
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                       }
                       
                       $result = Export-SubnetsToCsv -FilePath $saveDialog.FileName
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = $result
                           $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
                       }
                   }
               } catch {
                   if ($txtBlkImportStatus -ne $null) {
                       $txtBlkImportStatus.Text = "ERROR: $($_.Exception.Message)"
                       $txtBlkImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
                   }
               } finally {
                   $btnExportCsv.IsEnabled = $true
               }
           })
       }

       # --- ENTER KEY SUPPORT FOR LOOKUP TAB ---
       if ($txtIpLookup -ne $null) {
           $txtIpLookup.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       $ipToLookup = $txtIpLookup.Text.Trim()
                       Lookup-IpAddress -IpAddress $ipToLookup
                   }
               } catch {
               }
           })
       }

       # --- ENTER KEY SUPPORT FOR ADD TAB ---
       if ($txtIpSubnet -ne $null) {
           $txtIpSubnet.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtVlanId -ne $null) {
           $txtVlanId.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtVlanName -ne $null) {
           $txtVlanName.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       if ($txtSiteName -ne $null) {
           $txtSiteName.Add_KeyDown({
               param($sender, $e)
               try {
                   if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                       if ($btnAddEntry -ne $null) {
                           $btnAddEntry.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                       }
                   }
               } catch {
               }
           })
       }

       # Search functionality with debouncing
       if ($txtSearch -ne $null) {
           $txtSearch.Add_TextChanged({
               try {
                   $script:IPSearchTimer.Stop()
                   $script:IPSearchTimer.Start()
               } catch {
               }
           })
       }

       if ($btnClearSearch -ne $null) {
           $btnClearSearch.Add_Click({
               try {
                   if ($txtSearch -ne $null) { $txtSearch.Text = "" }
                   Update-SubnetDataGridWithSearch
               } catch {
               }
           })
       }

       if ($MainTabControl -ne $null) {
           $MainTabControl.Add_SelectionChanged({
               try {
                   $selectedTab = $MainTabControl.SelectedItem
                   
                   # Clear import/export status only when leaving Import tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Import") {
                       if ($txtBlkImportStatus -ne $null) {
                           $txtBlkImportStatus.Text = ""
                       }
                       if ($txtCsvFilePath -ne $null) {
                           $txtCsvFilePath.Text = ""
                       }
                   }
                   
                   # Clear add entry form AND status only when leaving Manage tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Manage Subnets") {
                       if ($txtIpSubnet -ne $null) { $txtIpSubnet.Text = "" }
                       if ($txtVlanId -ne $null) { $txtVlanId.Text = "" }
                       if ($txtVlanName -ne $null) { $txtVlanName.Text = "" }
                       if ($txtSiteName -ne $null) { $txtSiteName.Text = "" }
                   }

                   # Clear lookup form AND results only when leaving Lookup tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -ne "Lookup IP") {
                       if ($txtIpLookup -ne $null) { $txtIpLookup.Text = "" }
                       if ($grpLookupResults -ne $null) { $grpLookupResults.Visibility = "Collapsed" }
                       if ($txtBlkSearchedIp -ne $null) { $txtBlkSearchedIp.Text = "" }
                       if ($txtBlkMatchedSubnet -ne $null) { $txtBlkMatchedSubnet.Text = "" }
                       if ($txtBlkVlanId -ne $null) { $txtBlkVlanId.Text = "" }
                       if ($txtBlkVlanName -ne $null) { $txtBlkVlanName.Text = "" }
                       if ($txtBlkSiteName -ne $null) { $txtBlkSiteName.Text = "" }
                   }

                   # Show/hide status bar based on selected tab
                   if ($selectedTab -ne $null -and $selectedTab.Header -eq "Manage Subnets") {
                       if ($MainStatusBar -ne $null) {
                           $MainStatusBar.Visibility = [System.Windows.Visibility]::Visible
                           if ($dgSubnets.ItemsSource -ne $null) {
                               if ($txtStatusBarSubnets -ne $null) {
                                   $txtStatusBarSubnets.Text = "Total Subnets: $($dgSubnets.ItemsSource.Count)"
                               }
                           }
                       }
                   } else {
                       if ($MainStatusBar -ne $null) {
                           $MainStatusBar.Visibility = [System.Windows.Visibility]::Collapsed
                       }
                   }
               } catch {
               }
           })
       }      
   } catch {
   }
}