# ===================================================================
# IMPORT DEPENDENCIES
# ===================================================================

# Load device models and utilities first
. (Join-Path $PSScriptRoot "DeviceModels.ps1")
. (Join-Path $PSScriptRoot "SiteUtilities.ps1")
. (Join-Path $PSScriptRoot "SiteModels.ps1")

# Load WPF components if available
try {
    . (Join-Path $PSScriptRoot "WPFComponents.ps1")
} catch {
    Write-Verbose "WPF components not available - skipping PhoneNumberConverter"
}

# ===================================================================
# PHONE NUMBER CONVERTER CLASS
# ===================================================================

# Define XAML file path - Now in Core/Site/, need to go up two levels to UI
$xamlFile = Join-Path (Split-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -Parent) "UI" | Join-Path -ChildPath "NetworkManagement.xaml"

# ===================================================================
# SITE VALIDATION FUNCTIONS
# ===================================================================
function Validate-SiteBasicInfo {
    param(
        [SiteEntry]$Site,
        [object]$StatusControl = $null,
        [int]$ExcludeSiteID = -1  # For edit mode - exclude current site from duplicate checks
    )
    
    try {
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($Site.SiteCode)) {
            $errorMsg = "Site Code is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }

        if ([string]::IsNullOrWhiteSpace($Site.SiteSubnet)) {
            $errorMsg = "Site Subnet is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        if ([string]::IsNullOrWhiteSpace($Site.SiteName)) {
            $errorMsg = "Site Name is required and cannot be empty."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Validate subnet format
        if ($Site.SiteSubnet -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $octets = $Site.SiteSubnet.Split('.')
            $validOctets = $true
            foreach ($octet in $octets) {
                if ([int]$octet -lt 0 -or [int]$octet -gt 255) {
                    $validOctets = $false
                    break
                }
            }
            
            if (-not $validOctets) {
                $errorMsg = "Invalid subnet format. Each octet must be between 0-255."
                if ($StatusControl) {
                    $StatusControl.Text = $errorMsg
                    $StatusControl.Foreground = [System.Windows.Media.Brushes]::Red
                }
                throw $errorMsg
            }
        } else {
            $errorMsg = "Invalid subnet format. Please use format like: XXX.XX.XXX.XXX"
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Check for duplicates
        $allSites = $siteDataStore.GetAllEntries()
        
        # Check duplicate Site Code (exclude current site if editing)
        $duplicateSiteCode = $allSites | Where-Object { 
            $_.ID -ne $ExcludeSiteID -and $_.SiteCode -eq $Site.SiteCode 
        }
        if ($duplicateSiteCode) {
            $errorMsg = "Site code '$($Site.SiteCode)' already exists in another site."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # Check duplicate Site Subnet (exclude current site if editing)
        $duplicateSubnet = $allSites | Where-Object { 
            $_.ID -ne $ExcludeSiteID -and $_.SiteSubnet -eq $Site.SiteSubnet 
        }
        if ($duplicateSubnet) {
            $errorMsg = "Site subnet '$($Site.SiteSubnet)' already exists in another site."
            [StatusManager]::SetError($StatusControl, $errorMsg)
            throw $errorMsg
        }
        
        # If we get here, validation passed
        return $true
        
    } catch {
        # Re-throw the error for the calling function to handle
        throw $_
    }
}

# ===================================================================
# CENTRALIZED Centralized ComboBox Function
# ===================================================================
# Centralized function to set ComboBox selection by value or content

# ===================================================================
# UI DIALOG FUNCTIONS
# ===================================================================

# Show custom centered dialog with various button types and icons
function Show-CustomDialog {
    param(
        [string]$Message,
        [string]$Title,
        [string]$ButtonType = "OK",  # OK, YesNo, YesNoCancel
        [string]$Icon = "Information"  # Information, Warning, Error, Question
    )
    
    # Create a new window
    $dialog = New-Object System.Windows.Window
    $dialog.Title = $Title
    $dialog.Width = 400
    $dialog.Height = 200
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $mainWin
    $dialog.ResizeMode = "NoResize"
    $dialog.WindowStyle = "SingleBorderWindow"
    
    # Create the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "20"
    
    # Add row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = "*"
    $row2 = New-Object System.Windows.Controls.RowDefinition  
    $row2.Height = "Auto"
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    
    # Message text
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.Text = $Message
    $textBlock.TextWrapping = "Wrap"
    $textBlock.VerticalAlignment = "Center"
    $textBlock.HorizontalAlignment = "Center"
    $textBlock.FontSize = 12
    [System.Windows.Controls.Grid]::SetRow($textBlock, 0)
    $grid.Children.Add($textBlock)
    
    # Button panel
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Center"
    $buttonPanel.Margin = "0,20,0,0"
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
    
    $result = $null
    
    if ($ButtonType -eq "OK") {
        $okButton = New-Object System.Windows.Controls.Button
        $okButton.Content = "OK"
        $okButton.Width = 75
        $okButton.Height = 25
        $okButton.IsDefault = $true
        $okButton.Add_Click({
            $script:result = "OK"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        $buttonPanel.Children.Add($okButton)
    }
    elseif ($ButtonType -eq "YesNo") {
        $yesButton = New-Object System.Windows.Controls.Button
        $yesButton.Content = "Yes"
        $yesButton.Width = 75
        $yesButton.Height = 25
        $yesButton.Margin = "0,0,10,0"
        $yesButton.IsDefault = $true
        $yesButton.Add_Click({
            $script:result = "Yes"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        
        $noButton = New-Object System.Windows.Controls.Button
        $noButton.Content = "No"
        $noButton.Width = 75
        $noButton.Height = 25
        $noButton.IsCancel = $true
        $noButton.Add_Click({
            $script:result = "No"
            $dialog.DialogResult = $false
            $dialog.Close()
        })
        
        $buttonPanel.Children.Add($yesButton)
        $buttonPanel.Children.Add($noButton)
    }
    
    $grid.Children.Add($buttonPanel)
    $dialog.Content = $grid
    
    # Show dialog and return result
    $null = $dialog.ShowDialog()
    return $script:result
}

    # Show validation error message with status text update
    function Show-ValidationError {
        param(
            [string]$Message,
            [string]$Title = "Validation Error"
        )
        
        # Update status text safely
        try {
            if ($txtBlkSiteStatus) {
                $statusType = switch ($Title) {
                    "Success" { "Success" }
                    "Warning" { "Warning" }
                    default { "Error" }
                }
                [StatusManager]::SetStatus($txtBlkSiteStatus, $Message, $statusType)
            }
        } catch {
            # If status text update fails, just continue with dialog
        }
        
        # Show dialog
        Show-CustomDialog $Message $Title "OK" "Information"
    }

    # Create clickable text element for site details display
    function New-ClickableText {
        param(
            [string]$Value
        )
        
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.VerticalAlignment = 'Center'
        
        # Check for empty, null, or "(Not specified)" values
        if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq "(Not specified)") {
            $textBlock.Text = "(Not specified)"
            $textBlock.Foreground = [System.Windows.Media.Brushes]::Gray
            return $textBlock
        }
        
        # Make it clickable
        $textBlock.Text = $Value
        $textBlock.Cursor = [System.Windows.Input.Cursors]::Hand
        $textBlock.Foreground = [System.Windows.Media.Brushes]::Blue
        $textBlock.TextDecorations = [System.Windows.TextDecorations]::Underline
        $textBlock.ToolTip = "Click to copy: $Value"
        
        # Store original value as a property to avoid closure issues
        $textBlock | Add-Member -MemberType NoteProperty -Name "OriginalText" -Value $Value
        
        # Add click event to copy
        $textBlock.Add_MouseLeftButtonDown({
            param($sender, $e)
            try {
                $valueToCopy = $sender.OriginalText
                [System.Windows.Clipboard]::SetText($valueToCopy)
                
                # Simple feedback - change text briefly
                $sender.Text = "Copied!"
                $sender.Foreground = [System.Windows.Media.Brushes]::Green
                
                # Use DispatcherTimer for proper UI thread handling
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [System.TimeSpan]::FromMilliseconds(800)
                
                # Store reference to the textblock in timer's tag
                $timer.Tag = $sender
                
                $timer.Add_Tick({
                    param($timerSender, $timerArgs)
                    $textBlock = $timerSender.Tag
                    $textBlock.Text = $textBlock.OriginalText
                    $textBlock.Foreground = [System.Windows.Media.Brushes]::Blue
                    $timerSender.Stop()
                })
                $timer.Start()
                
            } catch {
            }
        })
        return $textBlock
    }


# Import device management functions
try {
    $deviceManagerPath = Join-Path $scriptPath "DeviceManager.ps1"    
    if (Test-Path $deviceManagerPath) {
        . $deviceManagerPath
    } 
    else {
        $errorMsg = "DeviceManager.ps1 not found at: $deviceManagerPath"
        Show-MessageBox $errorMsg "Module Error" "OK" "Error"
        exit 1
    }
}
catch {
    $errorMsg = "Failed to load DeviceManager.ps1: $_"
    Show-MessageBox $errorMsg "Module Error" "OK" "Error"
    exit 1
}


# ===================================================================
# FORM MANAGEMENT FUNCTIONS
# ===================================================================

# Clear all form fields and reset to default state
function Clear-SiteForm {
    # Clear all mapped fields using centralized field manager
    if ($script:FieldManager) {
        $script:FieldManager.ClearAllMappings()
    }
    
    # Reset devices using centralized manager
    if ($cmbSwitchCount) { $cmbSwitchCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('Switch', 0) }

    if ($cmbAPCount) { $cmbAPCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('AccessPoint', 0) }

    if ($cmbUPSCount) { $cmbUPSCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('UPS', 0) }

    if ($cmbCCTVCount) { $cmbCCTVCount.SelectedIndex = -1 }
    if ($script:DeviceManager) { $script:DeviceManager.UpdateDevicePanels('CCTV', 0) }
    
    # Reset main checkboxes
    if ($chkHasBackupCircuit) { $chkHasBackupCircuit.IsChecked = $false }
    if ($chkPrimaryHasModem) { $chkPrimaryHasModem.IsChecked = $false }
    if ($chkBackupHasModem) { $chkBackupHasModem.IsChecked = $false }
    
    # Hide conditional sections
    if ($grdBackupCircuit) { $grdBackupCircuit.Visibility = "Collapsed" }
    if ($stkPrimaryModem) { $stkPrimaryModem.Visibility = "Collapsed" }
    if ($stkBackupModem) { $stkBackupModem.Visibility = "Collapsed" }
}

# Collect all site data from form fields using centralized managers
function Get-SiteDataFromForm {
    $site = [SiteEntry]::new()
    
    # Get all mapped fields using centralized field manager
    $script:FieldManager.GetAllMappings($site)
    
    # Get device data using centralized device manager
    $site.SwitchCount = if ($cmbSwitchCount.SelectedItem) { [int]$cmbSwitchCount.SelectedItem.Content } else { 1 }
    $site.Switches = $script:DeviceManager.GetDeviceDataFromUI('Switch')

    $site.APCount = if ($cmbAPCount.SelectedItem) { [int]$cmbAPCount.SelectedItem.Content } else { 1 }
    $site.AccessPoints = $script:DeviceManager.GetDeviceDataFromUI('AccessPoint')

    $site.UPSCount = if ($cmbUPSCount.SelectedItem) { [int]$cmbUPSCount.SelectedItem.Content } else { 0 }
    $site.UPSDevices = $script:DeviceManager.GetDeviceDataFromUI('UPS')

    $site.CCTVCount = if ($cmbCCTVCount.SelectedItem) { [int]$cmbCCTVCount.SelectedItem.Content } else { 0 }
    $site.CCTVDevices = $script:DeviceManager.GetDeviceDataFromUI('CCTV')

    $site.PrinterCount = if ($cmbPrinterCount.SelectedItem) { [int]$cmbPrinterCount.SelectedItem.Content } else { 0 }
    $site.PrinterDevices = $script:DeviceManager.GetDeviceDataFromUI('Printer')
        
    # Get main checkboxes
    $site.HasBackupCircuit = $chkHasBackupCircuit.IsChecked
    
    return $site
}

# ===================================================================
# SITE MANAGEMENT FUNCTIONS
# ===================================================================

# Add new site with validation and duplicate checking
function Add-Site {
   try {
       # Get site data from form
       $site = Get-SiteDataFromForm

        # Use centralized validation
        try {
            Validate-SiteBasicInfo -Site $site -StatusControl $txtBlkSiteStatus
        } catch {
            Show-CustomDialog $_.Exception.Message "Validation Error" "OK" "Error"
            return $false
        }
       # Try to add the site
       try {
           $addResult = $siteDataStore.AddEntry($site)
           
           if ($addResult -eq $true) {
               Show-CustomDialog "Site '$($site.SiteCode)' added successfully!" "Success" "OK" "Information"
               Clear-SiteForm
               Update-DataGridWithSearch
               return $true
           }
       } catch {
           Show-CustomDialog "Error adding site: $($_.Exception.Message)" "Error" "OK" "Error"
           return $false
       }
   } catch {
       Show-CustomDialog "Error in Add-Site: $($_.Exception.Message)" "Error" "OK" "Error"
       return $false
   }
}

# ===================================================================
# DATA GRID MANAGEMENT FUNCTIONS
# ===================================================================

# Update DataGrid with search functionality and selection preservation
function Update-DataGridWithSearch {
    $searchTerm = $txtSearchSites.Text
    
    # Store current selection
    $selectedItems = @()
    if ($dgSites.SelectedItems.Count -gt 0) {
        foreach ($item in $dgSites.SelectedItems) {
            $selectedItems += $item.ID
        }
    }
    
    # Get all data from the data store
    $allData = $siteDataStore.GetAllEntries()
    
    # Filter data if search term exists
    if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
        $searchTerm = $searchTerm.Trim().ToLower()
        $allData = $allData | Where-Object {
            $_.SiteCode.ToLower().Contains($searchTerm) -or
            $_.SiteName.ToLower().Contains($searchTerm) -or
            $_.SiteSubnetCode.ToLower().Contains($searchTerm) -or
            $_.SiteAddress.ToLower().Contains($searchTerm) -or
            $_.MainContactName.ToLower().Contains($searchTerm) -or
            $_.Switch1IP.ToLower().Contains($searchTerm) -or
            $_.Switch1Name.ToLower().Contains($searchTerm) -or
            $_.FirewallIP.ToLower().Contains($searchTerm) -or
            $_.PrimaryVendor.ToLower().Contains($searchTerm)
        }
    }
    
    # Sort by ID numerically and update DataGrid
    $allData = $allData | Sort-Object -Property @{Expression={[int]$_.ID}; Ascending=$true}
    
    # Only update ItemsSource if data actually changed
    if ($dgSites.Items.Count -ne $allData.Count) {
        $dgSites.ItemsSource = @($allData)
        
        # Restore selection if items still exist
        if ($selectedItems.Count -gt 0) {
            $dgSites.SelectedItems.Clear()
            foreach ($item in $dgSites.Items) {
                if ($item.ID -in $selectedItems) {
                    $dgSites.SelectedItems.Add($item)
                }
            }
        }
    }
    
    # Update status bar
    $txtStatusBarSites.Text = "Total Sites: $($allData.Count)"
    $selectedCount = $dgSites.SelectedItems.Count
    if ($selectedCount -gt 0) {
        if ($selectedCount -eq 1) {
            $txtStatusBarSiteSelected.Text = "Selected: $($dgSites.SelectedItems[0].SiteCode) - $($dgSites.SelectedItems[0].SiteName)"
        } else {
            $txtStatusBarSiteSelected.Text = "Selected: $selectedCount sites"
        }
    } else {
        $txtStatusBarSiteSelected.Text = "Selected: None"
}}

# ===================================================================
# SITE LOOKUP AND DISPLAY FUNCTIONS
# ===================================================================

# Search for site by code or name and display details
function Lookup-Site {
    param([string]$SearchTerm)
    
    # Hide results at start
    $grpSiteLookupResults.Visibility = "Collapsed"
    
    if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
        Show-CustomDialog "Please enter a site code or name to search" "Input Required" "OK" "Warning"
        return
    }
    
    $searchTerm = $SearchTerm.Trim().ToLower()
    $allSites = $siteDataStore.GetAllEntries()
    
    $foundSite = $allSites | Where-Object {
        $_.SiteCode.ToLower().Contains($searchTerm) -or
        $_.SiteName.ToLower().Contains($searchTerm)
    } | Select-Object -First 1
    
    if ($foundSite) {
        Show-SiteDetails -Site $foundSite
        $grpSiteLookupResults.Visibility = "Visible"
    } else {
        Show-CustomDialog "Site '$SearchTerm' not found in the database." "Not Found" "OK" "Information"
    }
}

# Display comprehensive site details in lookup tab with three-column layout
function Show-SiteDetails {
    param([SiteEntry]$Site)
    
    try {
        $stkSiteDetails.Children.Clear()
        
        # Create main grid with three columns
        $mainGrid = New-Object System.Windows.Controls.Grid
        $mainGrid.Margin = "0"
        
        # Create three columns
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "1*"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "1*"
        $col3 = New-Object System.Windows.Controls.ColumnDefinition
        $col3.Width = "1*"
        $mainGrid.ColumnDefinitions.Add($col1)
        $mainGrid.ColumnDefinitions.Add($col2)
        $mainGrid.ColumnDefinitions.Add($col3)
        
        # Left Column StackPanel (Basic Info + Switches)
        $leftStack = New-Object System.Windows.Controls.StackPanel
        $leftStack.Margin = "0,0,10,0"
        [System.Windows.Controls.Grid]::SetColumn($leftStack, 0)
        
        # Middle Column StackPanel (Primary Circuit)
        $middleStack = New-Object System.Windows.Controls.StackPanel
        $middleStack.Margin = "5,0,5,0"
        [System.Windows.Controls.Grid]::SetColumn($middleStack, 1)
        
        # Right Column StackPanel (Backup Circuit + VLANs + Firewall)
        $rightStack = New-Object System.Windows.Controls.StackPanel
        $rightStack.Margin = "10,0,0,0"
        [System.Windows.Controls.Grid]::SetColumn($rightStack, 2)
        
        # === LEFT COLUMN CONTENT (Infrastructure & Basic Info) ===

        # Basic Info Section (keep as-is, stays first)
        $basicGroupBox = New-Object System.Windows.Controls.GroupBox
        $basicGroupBox.Header = "Basic Information"
        $basicGroupBox.Margin = "0,0,0,10"

        $basicGrid = New-Object System.Windows.Controls.Grid
        $basicGrid.Margin = "10"

        # Create columns for basic info
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $basicGrid.ColumnDefinitions.Add($col1)
        $basicGrid.ColumnDefinitions.Add($col2)

        # Create rows for basic info
        $basicInfoFields = @(
        @("Site Code:", $Site.SiteCode),
        @("Site Subnet:", $Site.SiteSubnet),
            @("Site Subnet Code:", $Site.SiteSubnetCode),
            @("Site Name:", $Site.SiteName),
            @("Site Address:", $Site.SiteAddress),
            @("Main Contact Name:", $Site.MainContactName),
            @("Main Contact Phone:", $Site.MainContactPhone),
            @("Second Contact Name:", $Site.SecondContactName),
            @("Second Contact Phone:", $Site.SecondContactPhone)
        )

        for ($i = 0; $i -lt $basicInfoFields.Count; $i++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $basicGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $basicInfoFields[$i][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $i)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $basicGrid.Children.Add($label)
            
            # Create clickable text
            $clickableText = New-ClickableText -Value $basicInfoFields[$i][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $basicGrid.Children.Add($clickableText)
            }

            $basicGroupBox.Content = $basicGrid
            $null = $leftStack.Children.Add($basicGroupBox)

            # Primary Circuit Section (moved to left column)
            $primaryGroupBox = New-Object System.Windows.Controls.GroupBox
            $primaryGroupBox.Header = "Primary Circuit"
            $primaryGroupBox.Margin = "0,0,0,10"

        # Check if primary circuit has any information
        $hasPrimaryInfo = (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.Vendor)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.CircuitType)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.CircuitID)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DownloadSpeed)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.UploadSpeed)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.IPAddress)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.SubnetMask)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DefaultGateway)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DNS1)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.DNS2)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterModel)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterName)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.RouterSN)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.PPPoEUsername)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.PPPoEPassword)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.ModemModel)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.ModemName)) -or
                        (-not [string]::IsNullOrWhiteSpace($Site.PrimaryCircuit.ModemSN))

        if (-not $hasPrimaryInfo) {
            $noPrimaryTextBlock = New-Object System.Windows.Controls.TextBlock
            $noPrimaryTextBlock.Text = "No primary circuit information configured for this site."
            $noPrimaryTextBlock.Margin = "10"
            $noPrimaryTextBlock.FontStyle = "Italic"
            $noPrimaryTextBlock.Foreground = "Gray"
            
            $primaryGroupBox.Content = $noPrimaryTextBlock
        } else {
            $primaryStack = New-Object System.Windows.Controls.StackPanel
            $primaryStack.Margin = "10"
            
            $primaryFields = @(
        @("Vendor:", $(if ($Site.PrimaryCircuit.Vendor) { $Site.PrimaryCircuit.Vendor } else { '(Not specified)' })),
        @("Circuit Type:", $(if ($Site.PrimaryCircuit.CircuitType) { $Site.PrimaryCircuit.CircuitType } else { '(Not specified)' })),
        @("Circuit ID:", $(if ($Site.PrimaryCircuit.CircuitID) { $Site.PrimaryCircuit.CircuitID } else { '(Not specified)' })),
        @("Download Speed:", $(if ($Site.PrimaryCircuit.DownloadSpeed) { $Site.PrimaryCircuit.DownloadSpeed } else { '(Not specified)' })),
        @("Upload Speed:", $(if ($Site.PrimaryCircuit.UploadSpeed) { $Site.PrimaryCircuit.UploadSpeed } else { '(Not specified)' })),
        @("IP Address:", $(if ($Site.PrimaryCircuit.IPAddress) { $Site.PrimaryCircuit.IPAddress } else { '(Not specified)' })),
        @("Subnet Mask:", $(if ($Site.PrimaryCircuit.SubnetMask) { $Site.PrimaryCircuit.SubnetMask } else { '(Not specified)' })),
        @("Default Gateway:", $(if ($Site.PrimaryCircuit.DefaultGateway) { $Site.PrimaryCircuit.DefaultGateway } else { '(Not specified)' })),
        @("DNS 1:", $(if ($Site.PrimaryCircuit.DNS1) { $Site.PrimaryCircuit.DNS1 } else { '(Not specified)' })),
        @("DNS 2:", $(if ($Site.PrimaryCircuit.DNS2) { $Site.PrimaryCircuit.DNS2 } else { '(Not specified)' })),
        @("Router Model:", $(if ($Site.PrimaryCircuit.RouterModel) { $Site.PrimaryCircuit.RouterModel } else { '(Not specified)' })),
        @("Router Name:", $(if ($Site.PrimaryCircuit.RouterName) { $Site.PrimaryCircuit.RouterName } else { '(Not specified)' })),
        @("Router Serial Number:", $(if ($Site.PrimaryCircuit.RouterSN) { $Site.PrimaryCircuit.RouterSN } else { '(Not specified)' }))
    )
    
    # Add GPON fields if applicable
    if ($Site.PrimaryCircuit.CircuitType -eq "GPON Fiber") {
        $primaryFields += @(
            @("PPPoE Username:", $(if ($Site.PrimaryCircuit.PPPoEUsername) { $Site.PrimaryCircuit.PPPoEUsername } else { '(Not specified)' })),
            @("PPPoE Password:", $(if ($Site.PrimaryCircuit.PPPoEPassword) { $Site.PrimaryCircuit.PPPoEPassword } else { '(Not specified)' }))
        )
    }
    
    # Add modem fields if applicable
    if ($Site.PrimaryCircuit.HasModem) {
        $primaryFields += @(
            @("Modem Model:", $(if ($Site.PrimaryCircuit.ModemModel) { $Site.PrimaryCircuit.ModemModel } else { '(Not specified)' })),
            @("Modem Name:", $(if ($Site.PrimaryCircuit.ModemName) { $Site.PrimaryCircuit.ModemName } else { '(Not specified)' })),
            @("Modem Serial Number:", $(if ($Site.PrimaryCircuit.ModemSN) { $Site.PrimaryCircuit.ModemSN } else { '(Not specified)' }))
        )
    }
    
    $primaryGrid = New-Object System.Windows.Controls.Grid
    $primaryGrid.Margin = "10"

    # Create columns
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = "Auto"
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = "*"
    $primaryGrid.ColumnDefinitions.Add($col1)
    $primaryGrid.ColumnDefinitions.Add($col2)

    for ($i = 0; $i -lt $primaryFields.Count; $i++) {
    $row = New-Object System.Windows.Controls.RowDefinition
    $row.Height = "Auto"
    $primaryGrid.RowDefinitions.Add($row)
    
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $primaryFields[$i][0]
    $label.FontWeight = "Bold"
    [System.Windows.Controls.Grid]::SetRow($label, $i)
    [System.Windows.Controls.Grid]::SetColumn($label, 0)
    $primaryGrid.Children.Add($label)
    
    # Create clickable text
    $clickableText = New-ClickableText -Value $primaryFields[$i][1]
    $clickableText.Margin = "10,5,0,5"
    [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
    $primaryGrid.Children.Add($clickableText)
    }

    $primaryGroupBox.Content = $primaryGrid
}

$null = $leftStack.Children.Add($primaryGroupBox)

# Backup Circuit Section (moved to left column)
if ($Site.HasBackupCircuit) {
    $backupGroupBox = New-Object System.Windows.Controls.GroupBox
    $backupGroupBox.Header = "Backup Circuit"
    $backupGroupBox.Margin = "0,0,0,10"
    
    $backupStack = New-Object System.Windows.Controls.StackPanel
    $backupStack.Margin = "10"
    
    $backupFields = @(
        @("Vendor:", $(if ($Site.BackupCircuit.Vendor) { $Site.BackupCircuit.Vendor } else { '(Not specified)' })),
        @("Circuit Type:", $(if ($Site.BackupCircuit.CircuitType) { $Site.BackupCircuit.CircuitType } else { '(Not specified)' })),
        @("Circuit ID:", $(if ($Site.BackupCircuit.CircuitID) { $Site.BackupCircuit.CircuitID } else { '(Not specified)' })),
        @("Download Speed:", $(if ($Site.BackupCircuit.DownloadSpeed) { $Site.BackupCircuit.DownloadSpeed } else { '(Not specified)' })),
        @("Upload Speed:", $(if ($Site.BackupCircuit.UploadSpeed) { $Site.BackupCircuit.UploadSpeed } else { '(Not specified)' })),
        @("IP Address:", $(if ($Site.BackupCircuit.IPAddress) { $Site.BackupCircuit.IPAddress } else { '(Not specified)' })),
        @("Subnet Mask:", $(if ($Site.BackupCircuit.SubnetMask) { $Site.BackupCircuit.SubnetMask } else { '(Not specified)' })),
        @("Default Gateway:", $(if ($Site.BackupCircuit.DefaultGateway) { $Site.BackupCircuit.DefaultGateway } else { '(Not specified)' })),
        @("DNS 1:", $(if ($Site.BackupCircuit.DNS1) { $Site.BackupCircuit.DNS1 } else { '(Not specified)' })),
        @("DNS 2:", $(if ($Site.BackupCircuit.DNS2) { $Site.BackupCircuit.DNS2 } else { '(Not specified)' })),
        @("Router Model:", $(if ($Site.BackupCircuit.RouterModel) { $Site.BackupCircuit.RouterModel } else { '(Not specified)' })),
        @("Router Name:", $(if ($Site.BackupCircuit.RouterName) { $Site.BackupCircuit.RouterName } else { '(Not specified)' })),
        @("Router Serial Number:", $(if ($Site.BackupCircuit.RouterSN) { $Site.BackupCircuit.RouterSN } else { '(Not specified)' }))
    )
    
    # Add GPON fields if applicable
    if ($Site.BackupCircuit.CircuitType -eq "GPON Fiber") {
        $backupFields += @(
            @("PPPoE Username:", $(if ($Site.BackupCircuit.PPPoEUsername) { $Site.BackupCircuit.PPPoEUsername } else { '(Not specified)' })),
            @("PPPoE Password:", $(if ($Site.BackupCircuit.PPPoEPassword) { $Site.BackupCircuit.PPPoEPassword } else { '(Not specified)' }))
        )
    }
    
    # Add modem fields if applicable
    if ($Site.BackupCircuit.HasModem) {
        $backupFields += @(
            @("Modem Model:", $(if ($Site.BackupCircuit.ModemModel) { $Site.BackupCircuit.ModemModel } else { '(Not specified)' })),
            @("Modem Name:", $(if ($Site.BackupCircuit.ModemName) { $Site.BackupCircuit.ModemName } else { '(Not specified)' })),
            @("Modem Serial Number:", $(if ($Site.BackupCircuit.ModemSN) { $Site.BackupCircuit.ModemSN } else { '(Not specified)' }))
        )
    }
    
    $backupGrid = New-Object System.Windows.Controls.Grid
    $backupGrid.Margin = "10"

    # Create columns
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = "Auto"
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = "*"
    $backupGrid.ColumnDefinitions.Add($col1)
    $backupGrid.ColumnDefinitions.Add($col2)

    for ($i = 0; $i -lt $backupFields.Count; $i++) {
    $row = New-Object System.Windows.Controls.RowDefinition
    $row.Height = "Auto"
    $backupGrid.RowDefinitions.Add($row)
    
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $backupFields[$i][0]
    $label.FontWeight = "Bold"
    [System.Windows.Controls.Grid]::SetRow($label, $i)
    [System.Windows.Controls.Grid]::SetColumn($label, 0)
    $backupGrid.Children.Add($label)
    
    # Create clickable text
    $clickableText = New-ClickableText -Value $backupFields[$i][1]
    $clickableText.Margin = "10,5,0,5"
    [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
    $backupGrid.Children.Add($clickableText)
    }

    $backupGroupBox.Content = $backupGrid
    $null = $leftStack.Children.Add($backupGroupBox)
} else {
    # Show that no backup circuit exists
    $noBackupGroupBox = New-Object System.Windows.Controls.GroupBox
    $noBackupGroupBox.Header = "Backup Circuit"
    $noBackupGroupBox.Margin = "0,0,0,10"
    
    $noBackupTextBlock = New-Object System.Windows.Controls.TextBlock
    $noBackupTextBlock.Text = "No backup circuit configured for this site."
    $noBackupTextBlock.Margin = "10"
    $noBackupTextBlock.FontStyle = "Italic"
    $noBackupTextBlock.Foreground = "Gray"
    
    $noBackupGroupBox.Content = $noBackupTextBlock
    $null = $leftStack.Children.Add($noBackupGroupBox)
}

# VLANs Section (moved to left column)
$vlanGroupBox = New-Object System.Windows.Controls.GroupBox
$vlanGroupBox.Header = "VLANs"
$vlanGroupBox.Margin = "0,0,0,10"

# Check if VLANs have any information
$hasVLANInfo = (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN100_Servers)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN101_NetworkDevices)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN102_UserDevices)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN103_UserDevices2)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN104_VOIP)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN105_WiFiCorp)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN106_WiFiBYOD)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN107_WiFiGuest)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN108_Spare)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN109_DMZ)) -or
            (-not [string]::IsNullOrWhiteSpace($Site.VLANs.VLAN110_CCTV))

if (-not $hasVLANInfo) {
    $noVLANTextBlock = New-Object System.Windows.Controls.TextBlock
    $noVLANTextBlock.Text = "No VLAN information configured for this site."
    $noVLANTextBlock.Margin = "10"
    $noVLANTextBlock.FontStyle = "Italic"
    $noVLANTextBlock.Foreground = "Gray"
    
    $vlanGroupBox.Content = $noVLANTextBlock
} else {
    $vlanStack = New-Object System.Windows.Controls.StackPanel
    $vlanStack.Margin = "10"
    
    $vlanFields = @(
        @("VLAN 100 - Servers:", $Site.VLANs.VLAN100_Servers),
        @("VLAN 101 - Network:", $Site.VLANs.VLAN101_NetworkDevices),
        @("VLAN 102 - User 1:", $Site.VLANs.VLAN102_UserDevices),
        @("VLAN 103 - User 2:", $Site.VLANs.VLAN103_UserDevices2),
        @("VLAN 104 - VOIP:", $Site.VLANs.VLAN104_VOIP),
        @("VLAN 105 - Wi-Fi Corp:", $Site.VLANs.VLAN105_WiFiCorp),
        @("VLAN 106 - Wi-Fi BYOD:", $Site.VLANs.VLAN106_WiFiBYOD),
        @("VLAN 107 - Wi-Fi Guest:", $Site.VLANs.VLAN107_WiFiGuest),
        @("VLAN 108 - Spare:", $Site.VLANs.VLAN108_Spare),
        @("VLAN 109 - DMZ:", $Site.VLANs.VLAN109_DMZ),
        @("VLAN 110 - CCTV:", $Site.VLANs.VLAN110_CCTV)
    )
    
    $vlanGrid = New-Object System.Windows.Controls.Grid
    $vlanGrid.Margin = "10"

    # Create columns
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = "Auto"
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = "*"
    $vlanGrid.ColumnDefinitions.Add($col1)
    $vlanGrid.ColumnDefinitions.Add($col2)

    for ($i = 0; $i -lt $vlanFields.Count; $i++) {
    $row = New-Object System.Windows.Controls.RowDefinition
    $row.Height = "Auto"
    $vlanGrid.RowDefinitions.Add($row)
    
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $vlanFields[$i][0]
    $label.FontWeight = "Bold"
    [System.Windows.Controls.Grid]::SetRow($label, $i)
    [System.Windows.Controls.Grid]::SetColumn($label, 0)
    $vlanGrid.Children.Add($label)
    
    # Create clickable text
    $clickableText = New-ClickableText -Value $vlanFields[$i][1]
    $clickableText.Margin = "10,5,0,5"
    [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
    $vlanGrid.Children.Add($clickableText)
    }

    $vlanGroupBox.Content = $vlanGrid
}

$null = $leftStack.Children.Add($vlanGroupBox)

        # === MIDDLE COLUMN CONTENT (Security & Network Devices) ===

# Firewall Section (moved to middle column)
$firewallGroupBox = New-Object System.Windows.Controls.GroupBox
$firewallGroupBox.Header = "Firewall"
$firewallGroupBox.Margin = "0,0,0,10"

# Check if firewall has any information
$hasFirewallInfo = (-not [string]::IsNullOrWhiteSpace($Site.FirewallIP)) -or
                (-not [string]::IsNullOrWhiteSpace($Site.FirewallName)) -or
                (-not [string]::IsNullOrWhiteSpace($Site.FirewallVersion)) -or
                (-not [string]::IsNullOrWhiteSpace($Site.FirewallSN))

if (-not $hasFirewallInfo) {
    $noFirewallTextBlock = New-Object System.Windows.Controls.TextBlock
    $noFirewallTextBlock.Text = "No firewall information configured for this site."
    $noFirewallTextBlock.Margin = "10"
    $noFirewallTextBlock.FontStyle = "Italic"
    $noFirewallTextBlock.Foreground = "Gray"
    
    $firewallGroupBox.Content = $noFirewallTextBlock
} else {
    $firewallStack = New-Object System.Windows.Controls.StackPanel
    $firewallStack.Margin = "10"
    
    $firewallFields = @(
        @("Management IP:", $(if ($Site.FirewallIP) { $Site.FirewallIP } else { '(Not specified)' })),
        @("Name:", $(if ($Site.FirewallName) { $Site.FirewallName } else { '(Not specified)' })),
        @("Version:", $(if ($Site.FirewallVersion) { $Site.FirewallVersion } else { '(Not specified)' })),
        @("Serial Number:", $(if ($Site.FirewallSN) { $Site.FirewallSN } else { '(Not specified)' }))
    )
    
    $firewallGrid = New-Object System.Windows.Controls.Grid
    $firewallGrid.Margin = "10"

    # Create columns
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = "Auto"
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = "*"
    $firewallGrid.ColumnDefinitions.Add($col1)
    $firewallGrid.ColumnDefinitions.Add($col2)

    for ($i = 0; $i -lt $firewallFields.Count; $i++) {
    $row = New-Object System.Windows.Controls.RowDefinition
    $row.Height = "Auto"
    $firewallGrid.RowDefinitions.Add($row)
    
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $firewallFields[$i][0]
    $label.FontWeight = "Bold"
    [System.Windows.Controls.Grid]::SetRow($label, $i)
    [System.Windows.Controls.Grid]::SetColumn($label, 0)
    $firewallGrid.Children.Add($label)
    
    # Create clickable text
    $clickableText = New-ClickableText -Value $firewallFields[$i][1]
    $clickableText.Margin = "10,5,0,5"
    [System.Windows.Controls.Grid]::SetRow($clickableText, $i)
    [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
    $firewallGrid.Children.Add($clickableText)
    }

    $firewallGroupBox.Content = $firewallGrid
}

$null = $middleStack.Children.Add($firewallGroupBox)

# Switches Section (moved to middle column)
$switchesGroupBox = New-Object System.Windows.Controls.GroupBox
# Count actual switches with data
$actualSwitchCount = 0
foreach ($switch in $Site.Switches) {
    if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP) -or 
        -not [string]::IsNullOrWhiteSpace($switch.Name) -or 
        -not [string]::IsNullOrWhiteSpace($switch.AssetTag) -or 
        -not [string]::IsNullOrWhiteSpace($switch.Version) -or 
        -not [string]::IsNullOrWhiteSpace($switch.SerialNumber)) {
        $actualSwitchCount++
    }
}
$switchesGroupBox.Header = "Switches ($actualSwitchCount total)"
$switchesGroupBox.Margin = "0,0,0,10"

$switchesStack = New-Object System.Windows.Controls.StackPanel
$switchesStack.Margin = "10"

# Handle case where there are no switches OR all switches are empty
if ($Site.Switches.Count -eq 0) {
    $noSwitchTextBlock = New-Object System.Windows.Controls.TextBlock
    $noSwitchTextBlock.Text = "No switches configured for this site."
    $noSwitchTextBlock.FontStyle = "Italic"
    $noSwitchTextBlock.Foreground = "Gray"
    $noSwitchTextBlock.Margin = "0,5,0,5"
    $null = $switchesStack.Children.Add($noSwitchTextBlock)
} else {
    # Check if all switches are empty
    $hasValidSwitches = $false
    foreach ($switch in $Site.Switches) {
        if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP) -or 
            -not [string]::IsNullOrWhiteSpace($switch.Name) -or 
            -not [string]::IsNullOrWhiteSpace($switch.AssetTag) -or 
            -not [string]::IsNullOrWhiteSpace($switch.Version) -or 
            -not [string]::IsNullOrWhiteSpace($switch.SerialNumber)) {
            $hasValidSwitches = $true
            break
        }
    }
    
    if (-not $hasValidSwitches) {
        $noSwitchTextBlock = New-Object System.Windows.Controls.TextBlock
        $noSwitchTextBlock.Text = "No switch information configured for this site."
        $noSwitchTextBlock.FontStyle = "Italic"
        $noSwitchTextBlock.Foreground = "Gray"
        $noSwitchTextBlock.Margin = "0,5,0,5"
        $null = $switchesStack.Children.Add($noSwitchTextBlock)
    } else {
        for ($i = 0; $i -lt $Site.Switches.Count; $i++) {
            $switch = $Site.Switches[$i]
            $switchTextBlock = New-Object System.Windows.Controls.TextBlock
            $switchTextBlock.FontWeight = "Bold"
            $switchTextBlock.Text = "Switch $($i+1):"
            $switchTextBlock.Margin = "0,5,0,2"
            $null = $switchesStack.Children.Add($switchTextBlock)
            
            # Switch details with bold labels
            $switchDetails = @(
                @("Management IP:", $(if ($switch.ManagementIP) { $switch.ManagementIP } else { '(Not specified)' })),
                @("Name:", $(if ($switch.Name) { $switch.Name } else { '(Not specified)' })),
                @("Asset Tag:", $(if ($switch.AssetTag) { $switch.AssetTag } else { '(Not specified)' })),
                @("Version:", $(if ($switch.Version) { $switch.Version } else { '(Not specified)' })),
                @("Serial Number:", $(if ($switch.SerialNumber) { $switch.SerialNumber } else { '(Not specified)' }))
            )
            
            $switchGrid = New-Object System.Windows.Controls.Grid
            $switchGrid.Margin = "0,0,0,5"

            # Create columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = "*"
            $switchGrid.ColumnDefinitions.Add($col1)
            $switchGrid.ColumnDefinitions.Add($col2)

            for ($j = 0; $j -lt $switchDetails.Count; $j++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $switchGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $switchDetails[$j][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $j)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $switchGrid.Children.Add($label)
            
            # Create clickable text
            $clickableText = New-ClickableText -Value $switchDetails[$j][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $switchGrid.Children.Add($clickableText)
            }

            $null = $switchesStack.Children.Add($switchGrid)
        }
    }
}

$switchesGroupBox.Content = $switchesStack
$null = $middleStack.Children.Add($switchesGroupBox)

# Access Points Section (moved to middle column)
$apGroupBox = New-Object System.Windows.Controls.GroupBox
# Count actual APs with data
$actualAPCount = 0
foreach ($ap in $Site.AccessPoints) {
    if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP) -or 
        -not [string]::IsNullOrWhiteSpace($ap.Name) -or 
        -not [string]::IsNullOrWhiteSpace($ap.AssetTag) -or 
        -not [string]::IsNullOrWhiteSpace($ap.Version) -or 
        -not [string]::IsNullOrWhiteSpace($ap.SerialNumber)) {
        $actualAPCount++
    }
}
$apGroupBox.Header = "Access Points ($actualAPCount total)"
$apGroupBox.Margin = "0,0,0,10"

$apStack = New-Object System.Windows.Controls.StackPanel
$apStack.Margin = "10"

# Handle case where there are no access points OR all APs are empty
if ($Site.AccessPoints.Count -eq 0) {
    $noApTextBlock = New-Object System.Windows.Controls.TextBlock
    $noApTextBlock.Text = "No access points configured for this site."
    $noApTextBlock.FontStyle = "Italic"
    $noApTextBlock.Foreground = "Gray"
    $noApTextBlock.Margin = "0,5,0,5"
    $null = $apStack.Children.Add($noApTextBlock)
} else {
    # Check if all APs are empty
    $hasValidAPs = $false
    foreach ($ap in $Site.AccessPoints) {
        if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP) -or 
            -not [string]::IsNullOrWhiteSpace($ap.Name) -or 
            -not [string]::IsNullOrWhiteSpace($ap.AssetTag) -or 
            -not [string]::IsNullOrWhiteSpace($ap.Version) -or 
            -not [string]::IsNullOrWhiteSpace($ap.SerialNumber)) {
            $hasValidAPs = $true
            break
        }
    }
    
    if (-not $hasValidAPs) {
        $noApTextBlock = New-Object System.Windows.Controls.TextBlock
        $noApTextBlock.Text = "No access point information configured for this site."
        $noApTextBlock.FontStyle = "Italic"
        $noApTextBlock.Foreground = "Gray"
        $noApTextBlock.Margin = "0,5,0,5"
        $null = $apStack.Children.Add($noApTextBlock)
    } else {
        for ($i = 0; $i -lt $Site.AccessPoints.Count; $i++) {
            $ap = $Site.AccessPoints[$i]
            $apTextBlock = New-Object System.Windows.Controls.TextBlock
            $apTextBlock.FontWeight = "Bold"
            $apTextBlock.Text = "Access Point $($i+1):"
            $apTextBlock.Margin = "0,5,0,2"
            $null = $apStack.Children.Add($apTextBlock)
            
            # AP details with grid layout for nice formatting
            $apGrid = New-Object System.Windows.Controls.Grid
            $apGrid.Margin = "0,0,0,5"

            # Create columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = "*"
            $apGrid.ColumnDefinitions.Add($col1)
            $apGrid.ColumnDefinitions.Add($col2)

            $apDetails = @(
                @("Management IP:", $(if ($ap.ManagementIP) { $ap.ManagementIP } else { '(Not specified)' })),
                @("Name:", $(if ($ap.Name) { $ap.Name } else { '(Not specified)' })),
                @("Asset Tag:", $(if ($ap.AssetTag) { $ap.AssetTag } else { '(Not specified)' })),
                @("Version:", $(if ($ap.Version) { $ap.Version } else { '(Not specified)' })),
                @("Serial Number:", $(if ($ap.SerialNumber) { $ap.SerialNumber } else { '(Not specified)' }))
            )

            for ($j = 0; $j -lt $apDetails.Count; $j++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $apGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $apDetails[$j][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $j)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $apGrid.Children.Add($label)
            
            # Create clickable text
            $clickableText = New-ClickableText -Value $apDetails[$j][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $apGrid.Children.Add($clickableText)
            }
            
            $null = $apStack.Children.Add($apGrid)
        }
    }
}

$apGroupBox.Content = $apStack
$null = $middleStack.Children.Add($apGroupBox)

# === RIGHT COLUMN CONTENT (Monitoring & Surveillance) ===

# UPS Section (moved to right column)
$upsGroupBox = New-Object System.Windows.Controls.GroupBox
# Count actual UPS with data
$actualUPSCount = 0
foreach ($ups in $Site.UPSDevices) {
    if (-not [string]::IsNullOrWhiteSpace($ups.ManagementIP) -or 
        -not [string]::IsNullOrWhiteSpace($ups.Name)) {
        $actualUPSCount++
    }
}
$upsGroupBox.Header = "UPS ($actualUPSCount total)"
$upsGroupBox.Margin = "0,0,0,10"

$upsStack = New-Object System.Windows.Controls.StackPanel
$upsStack.Margin = "10"

# Handle case where there are no UPS devices
if ($Site.UPSDevices.Count -eq 0) {
    $noUpsTextBlock = New-Object System.Windows.Controls.TextBlock
    $noUpsTextBlock.Text = "No UPS devices configured for this site."
    $noUpsTextBlock.FontStyle = "Italic"
    $noUpsTextBlock.Foreground = "Gray"
    $noUpsTextBlock.Margin = "0,5,0,5"
    $null = $upsStack.Children.Add($noUpsTextBlock)
} else {
    for ($i = 0; $i -lt $Site.UPSDevices.Count; $i++) {
        $ups = $Site.UPSDevices[$i]
        $upsTextBlock = New-Object System.Windows.Controls.TextBlock
        $upsTextBlock.FontWeight = "Bold"
        $upsTextBlock.Text = "UPS $($i+1):"
        $upsTextBlock.Margin = "0,5,0,2"
        $null = $upsStack.Children.Add($upsTextBlock)
        
        # UPS details with grid layout - only IP and Name now
        $upsGrid = New-Object System.Windows.Controls.Grid
        $upsGrid.Margin = "0,0,0,5"

        # Create columns
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $upsGrid.ColumnDefinitions.Add($col1)
        $upsGrid.ColumnDefinitions.Add($col2)

        $upsDetails = @(
            @("Management IP:", $(if ($ups.ManagementIP) { $ups.ManagementIP } else { '(Not specified)' })),
            @("Name:", $(if ($ups.Name) { $ups.Name } else { '(Not specified)' }))
        )

        for ($j = 0; $j -lt $upsDetails.Count; $j++) {
        $row = New-Object System.Windows.Controls.RowDefinition
        $row.Height = "Auto"
        $upsGrid.RowDefinitions.Add($row)
        
        $label = New-Object System.Windows.Controls.Label
        $label.Content = $upsDetails[$j][0]
        $label.FontWeight = "Bold"
        [System.Windows.Controls.Grid]::SetRow($label, $j)
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        $upsGrid.Children.Add($label)
        
        # Create clickable text
        $clickableText = New-ClickableText -Value $upsDetails[$j][1]
        $clickableText.Margin = "10,5,0,5"
        [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
        [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
        $upsGrid.Children.Add($clickableText)
        }
        
        $null = $upsStack.Children.Add($upsGrid)
    }
}

$upsGroupBox.Content = $upsStack
$null = $rightStack.Children.Add($upsGroupBox)

# CCTV Section (moved to right column)
$cctvGroupBox = New-Object System.Windows.Controls.GroupBox
# Count actual CCTV with data
$actualCCTVCount = 0
foreach ($cctv in $Site.CCTVDevices) {
    if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP) -or 
        -not [string]::IsNullOrWhiteSpace($cctv.Name) -or 
        -not [string]::IsNullOrWhiteSpace($cctv.SerialNumber)) {
        $actualCCTVCount++
    }
}
$cctvGroupBox.Header = "CCTV Cameras ($actualCCTVCount total)"

# Printer Section (moved to right column)
$printerGroupBox = New-Object System.Windows.Controls.GroupBox
# Count actual Printers with data
$actualPrinterCount = 0
foreach ($printer in $Site.PrinterDevices) {
    if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP) -or 
        -not [string]::IsNullOrWhiteSpace($printer.Name) -or 
        -not [string]::IsNullOrWhiteSpace($printer.Model) -or
        -not [string]::IsNullOrWhiteSpace($printer.SerialNumber)) {
        $actualPrinterCount++
    }
}
$printerGroupBox.Header = "Printers ($actualPrinterCount total)"
$printerGroupBox.Margin = "0,0,0,10"

$printerStack = New-Object System.Windows.Controls.StackPanel
$printerStack.Margin = "10"

# Handle case where there are no Printer devices
if ($Site.PrinterDevices.Count -eq 0) {
    $noPrinterTextBlock = New-Object System.Windows.Controls.TextBlock
    $noPrinterTextBlock.Text = "No printers configured for this site."
    $noPrinterTextBlock.FontStyle = "Italic"
    $noPrinterTextBlock.Foreground = "Gray"
    $noPrinterTextBlock.Margin = "0,5,0,5"
    $null = $printerStack.Children.Add($noPrinterTextBlock)
} else {
    # Check if all Printers are empty
    $hasValidPrinters = $false
    foreach ($printer in $Site.PrinterDevices) {
        if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP) -or 
            -not [string]::IsNullOrWhiteSpace($printer.Name) -or 
            -not [string]::IsNullOrWhiteSpace($printer.Model) -or
            -not [string]::IsNullOrWhiteSpace($printer.SerialNumber)) {
            $hasValidPrinters = $true
            break
        }
    }
    
    if (-not $hasValidPrinters) {
        $noPrinterTextBlock = New-Object System.Windows.Controls.TextBlock
        $noPrinterTextBlock.Text = "No printer information configured for this site."
        $noPrinterTextBlock.FontStyle = "Italic"
        $noPrinterTextBlock.Foreground = "Gray"
        $noPrinterTextBlock.Margin = "0,5,0,5"
        $null = $printerStack.Children.Add($noPrinterTextBlock)
    } else {
        for ($i = 0; $i -lt $Site.PrinterDevices.Count; $i++) {
            $printer = $Site.PrinterDevices[$i]
            $printerTextBlock = New-Object System.Windows.Controls.TextBlock
            $printerTextBlock.FontWeight = "Bold"
            $printerTextBlock.Text = "Printer $($i+1):"
            $printerTextBlock.Margin = "0,5,0,2"
            $null = $printerStack.Children.Add($printerTextBlock)
            
            # Printer details with grid layout
            $printerGrid = New-Object System.Windows.Controls.Grid
            $printerGrid.Margin = "0,0,0,5"

            # Create columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = "*"
            $printerGrid.ColumnDefinitions.Add($col1)
            $printerGrid.ColumnDefinitions.Add($col2)

            $printerDetails = @(
                @("Management IP:", $(if ($printer.ManagementIP) { $printer.ManagementIP } else { '(Not specified)' })),
                @("Name:", $(if ($printer.Name) { $printer.Name } else { '(Not specified)' })),
                @("Model:", $(if ($printer.Model) { $printer.Model } else { '(Not specified)' })),
                @("Serial Number:", $(if ($printer.SerialNumber) { $printer.SerialNumber } else { '(Not specified)' }))
            )

            for ($j = 0; $j -lt $printerDetails.Count; $j++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $printerGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $printerDetails[$j][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $j)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $printerGrid.Children.Add($label)
            
            # Create clickable text
            $clickableText = New-ClickableText -Value $printerDetails[$j][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $printerGrid.Children.Add($clickableText)
            }
            
            $null = $printerStack.Children.Add($printerGrid)
        }
    }
}

$printerGroupBox.Content = $printerStack
$null = $rightStack.Children.Add($printerGroupBox)
$cctvGroupBox.Margin = "0,0,0,10"

$cctvStack = New-Object System.Windows.Controls.StackPanel
$cctvStack.Margin = "10"

# Handle case where there are no CCTV devices
if ($Site.CCTVDevices.Count -eq 0) {
    $noCctvTextBlock = New-Object System.Windows.Controls.TextBlock
    $noCctvTextBlock.Text = "No CCTV cameras configured for this site."
    $noCctvTextBlock.FontStyle = "Italic"
    $noCctvTextBlock.Foreground = "Gray"
    $noCctvTextBlock.Margin = "0,5,0,5"
    $null = $cctvStack.Children.Add($noCctvTextBlock)
} else {
    # Check if all CCTV cameras are empty
    $hasValidCCTV = $false
    foreach ($cctv in $Site.CCTVDevices) {
        if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP) -or 
            -not [string]::IsNullOrWhiteSpace($cctv.Name) -or 
            -not [string]::IsNullOrWhiteSpace($cctv.SerialNumber)) {
            $hasValidCCTV = $true
            break
        }
    }
    
    if (-not $hasValidCCTV) {
        $noCctvTextBlock = New-Object System.Windows.Controls.TextBlock
        $noCctvTextBlock.Text = "No CCTV camera information configured for this site."
        $noCctvTextBlock.FontStyle = "Italic"
        $noCctvTextBlock.Foreground = "Gray"
        $noCctvTextBlock.Margin = "0,5,0,5"
        $null = $cctvStack.Children.Add($noCctvTextBlock)
    } else {
        for ($i = 0; $i -lt $Site.CCTVDevices.Count; $i++) {
            $cctv = $Site.CCTVDevices[$i]
            $cctvTextBlock = New-Object System.Windows.Controls.TextBlock
            $cctvTextBlock.FontWeight = "Bold"
            $cctvTextBlock.Text = "Camera $($i+1):"
            $cctvTextBlock.Margin = "0,5,0,2"
            $null = $cctvStack.Children.Add($cctvTextBlock)
            
            # CCTV details with grid layout
            $cctvGrid = New-Object System.Windows.Controls.Grid
            $cctvGrid.Margin = "0,0,0,5"

            # Create columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = "Auto"
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = "*"
            $cctvGrid.ColumnDefinitions.Add($col1)
            $cctvGrid.ColumnDefinitions.Add($col2)

            $cctvDetails = @(
                @("Management IP:", $(if ($cctv.ManagementIP) { $cctv.ManagementIP } else { '(Not specified)' })),
                @("Name:", $(if ($cctv.Name) { $cctv.Name } else { '(Not specified)' })),
                @("Serial Number:", $(if ($cctv.SerialNumber) { $cctv.SerialNumber } else { '(Not specified)' }))
            )

            for ($j = 0; $j -lt $cctvDetails.Count; $j++) {
            $row = New-Object System.Windows.Controls.RowDefinition
            $row.Height = "Auto"
            $cctvGrid.RowDefinitions.Add($row)
            
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $cctvDetails[$j][0]
            $label.FontWeight = "Bold"
            [System.Windows.Controls.Grid]::SetRow($label, $j)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            $cctvGrid.Children.Add($label)
            
            # Create clickable text
            $clickableText = New-ClickableText -Value $cctvDetails[$j][1]
            $clickableText.Margin = "10,5,0,5"
            [System.Windows.Controls.Grid]::SetRow($clickableText, $j)
            [System.Windows.Controls.Grid]::SetColumn($clickableText, 1)
            $cctvGrid.Children.Add($clickableText)
            }
            
            $null = $cctvStack.Children.Add($cctvGrid)
        }
    }
}

$cctvGroupBox.Content = $cctvStack
$null = $rightStack.Children.Add($cctvGroupBox)

# Add all three columns to main grid
$mainGrid.Children.Add($leftStack)
$mainGrid.Children.Add($middleStack)
$mainGrid.Children.Add($rightStack)

# Add main grid to the site details stack
$null = $stkSiteDetails.Children.Add($mainGrid)

    }
    catch {
        Show-MessageBox "Error displaying site details: $_" "Display Error" "OK" "Error"
    }
}

# ===================================================================
# EVENT HANDLER FUNCTIONS
# ===================================================================

# Generic device count change handler for all device types
function Handle-DeviceCountChanged {
    param(
        [string]$DeviceType,
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    if ($ComboBox.SelectedItem) {
        $count = [int]$ComboBox.SelectedItem.Content
        $script:DeviceManager.UpdateDevicePanels($DeviceType, $count)
        Write-Host "DEBUG: About to update panels for DeviceType: '$DeviceType', count: $count"
        
        # Auto-populate names and IPs if site code/subnet exists
        if (-not [string]::IsNullOrWhiteSpace($txtSiteCode.Text)) {
            $script:DeviceManager.UpdateDeviceNamesFromSiteCode($DeviceType, $txtSiteCode.Text)
        }
        if (-not [string]::IsNullOrWhiteSpace($txtSiteSubnet.Text)) {
            if ($txtSiteSubnet.Text -match '^(\d+\.\d+)\.') {
                $script:DeviceManager.UpdateDeviceIPsFromSubnet($DeviceType, $matches[1])
            }
        }
    }
}

# ===================================================================
# XAML LOADING AND UI INITIALIZATION
# ===================================================================

# When dot-sourced from Main.ps1, we need to get the actual Site.ps1 directory
$scriptPath = $PSScriptRoot
if (-not $scriptPath) {
    # Fallback: assume we're in Site Network Identifier folder
    $scriptPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Site Network Identifier"
}

# Validate XAML file exists and is readable
if (-not (Test-Path $xamlFile)) {
    Show-MessageBox "XAML file not found: $xamlFile" "File Error" "OK" "Error"
    exit
}

# Load and parse XAML
try {
    $xaml = Get-Content $xamlFile -Raw
    $xml = [xml]$xaml
}
catch {
    Show-MessageBox "Error loading XAML: $_" "XAML Error" "OK" "Error"
    exit
}

# Read XAML and create GUI elements
try {
    # Create the window first without loading XAML content
    $reader = New-Object System.Xml.XmlNodeReader $xml
    
    # Add the phone formatter resource BEFORE loading XAML
    $phoneConverter = [PhoneNumberConverter]::new()
    
    # Create a resource dictionary and add our converter
    $resourceDict = New-Object System.Windows.ResourceDictionary
    $resourceDict.Add("PhoneFormatter", $phoneConverter)
    
    # Load the window with resources
    $mainWin = [Windows.Markup.XamlReader]::Load($reader)
    $mainWin.Resources = $resourceDict
}
catch {
    Show-MessageBox "Error creating window: $_" "Window Creation Error" "OK" "Error"
    exit
}

# ===================================================================
# UI CONTROL REFERENCES
# ===================================================================

# Get GUI elements by Name - Basic Info
$txtSiteCode = $mainWin.FindName("txtSiteCode")
$txtSiteSubnetCode = $mainWin.FindName("txtSiteSubnetCode")
$txtSiteSubnet = $mainWin.FindName("txtSiteSubnet")
$txtSiteName = $mainWin.FindName("txtSiteNameManage")
$txtSiteAddress = $mainWin.FindName("txtSiteAddress")
$txtMainContactName = $mainWin.FindName("txtMainContactName")
$txtMainContactPhone = $mainWin.FindName("txtMainContactPhone")
$txtSecondContactName = $mainWin.FindName("txtSecondContactName")
$txtSecondContactPhone = $mainWin.FindName("txtSecondContactPhone")

# Network Equipment
$cmbSwitchCount = $mainWin.FindName("cmbSwitchCount")
$stkSwitches = $mainWin.FindName("stkSwitches")
$txtFirewallIP = $mainWin.FindName("txtFirewallIP")
$txtFirewallName = $mainWin.FindName("txtFirewallName")
$txtFirewallVersion = $mainWin.FindName("txtFirewallVersion")
$txtFirewallSN = $mainWin.FindName("txtFirewallSN")

# Primary Circuit
$txtPrimaryVendor = $mainWin.FindName("txtPrimaryVendor")
$cmbPrimaryCircuitType = $mainWin.FindName("cmbPrimaryCircuitType")
$txtPrimaryCircuitID = $mainWin.FindName("txtPrimaryCircuitID")
$txtPrimaryDownloadSpeed = $mainWin.FindName("txtPrimaryDownloadSpeed")
$txtPrimaryUploadSpeed = $mainWin.FindName("txtPrimaryUploadSpeed")
$txtPrimaryIPAddress = $mainWin.FindName("txtPrimaryIPAddress")
$txtPrimarySubnetMask = $mainWin.FindName("txtPrimarySubnetMask")
$txtPrimaryDefaultGateway = $mainWin.FindName("txtPrimaryDefaultGateway")
$txtPrimaryDNS1 = $mainWin.FindName("txtPrimaryDNS1")
$txtPrimaryDNS2 = $mainWin.FindName("txtPrimaryDNS2")
$txtPrimaryRouterModel = $mainWin.FindName("txtPrimaryRouterModel")
$txtPrimaryRouterName = $mainWin.FindName("txtPrimaryRouterName")
$txtPrimaryRouterSN = $mainWin.FindName("txtPrimaryRouterSN")
$txtPrimaryPPPoEUsername = $mainWin.FindName("txtPrimaryPPPoEUsername")
$txtPrimaryPPPoEPassword = $mainWin.FindName("txtPrimaryPPPoEPassword")
$chkPrimaryHasModem = $mainWin.FindName("chkPrimaryHasModem")
$stkPrimaryModem = $mainWin.FindName("stkPrimaryModem")
$txtPrimaryModemModel = $mainWin.FindName("txtPrimaryModemModel")
$txtPrimaryModemName = $mainWin.FindName("txtPrimaryModemName")
$txtPrimaryModemSN = $mainWin.FindName("txtPrimaryModemSN")

# Backup Circuit
$chkHasBackupCircuit = $mainWin.FindName("chkHasBackupCircuit")
$grdBackupCircuit = $mainWin.FindName("grdBackupCircuit")
$txtBackupVendor = $mainWin.FindName("txtBackupVendor")
$cmbBackupCircuitType = $mainWin.FindName("cmbBackupCircuitType")
$txtBackupCircuitID = $mainWin.FindName("txtBackupCircuitID")
$txtBackupDownloadSpeed = $mainWin.FindName("txtBackupDownloadSpeed")
$txtBackupUploadSpeed = $mainWin.FindName("txtBackupUploadSpeed")
$txtBackupIPAddress = $mainWin.FindName("txtBackupIPAddress")
$txtBackupSubnetMask = $mainWin.FindName("txtBackupSubnetMask")
$txtBackupDefaultGateway = $mainWin.FindName("txtBackupDefaultGateway")
$txtBackupDNS1 = $mainWin.FindName("txtBackupDNS1")
$txtBackupDNS2 = $mainWin.FindName("txtBackupDNS2")
$txtBackupRouterModel = $mainWin.FindName("txtBackupRouterModel")
$txtBackupRouterName = $mainWin.FindName("txtBackupRouterName")
$txtBackupRouterSN = $mainWin.FindName("txtBackupRouterSN")
$txtBackupPPPoEUsername = $mainWin.FindName("txtBackupPPPoEUsername")
$txtBackupPPPoEPassword = $mainWin.FindName("txtBackupPPPoEPassword")
$chkBackupHasModem = $mainWin.FindName("chkBackupHasModem")
$stkBackupModem = $mainWin.FindName("stkBackupModem")
$txtBackupModemModel = $mainWin.FindName("txtBackupModemModel")
$txtBackupModemName = $mainWin.FindName("txtBackupModemName")
$txtBackupModemSN = $mainWin.FindName("txtBackupModemSN")

# VLANs
$txtVlan100 = $mainWin.FindName("txtVlan100")
$txtVlan101 = $mainWin.FindName("txtVlan101")
$txtVlan102 = $mainWin.FindName("txtVlan102")
$txtVlan103 = $mainWin.FindName("txtVlan103")
$txtVlan104 = $mainWin.FindName("txtVlan104")
$txtVlan105 = $mainWin.FindName("txtVlan105")
$txtVlan106 = $mainWin.FindName("txtVlan106")
$txtVlan107 = $mainWin.FindName("txtVlan107")
$txtVlan108 = $mainWin.FindName("txtVlan108")
$txtVlan109 = $mainWin.FindName("txtVlan109")
$txtVlan110 = $mainWin.FindName("txtVlan110")

# Access Points
$cmbAPCount = $mainWin.FindName("cmbAPCount")
$stkAccessPoints = $mainWin.FindName("stkAccessPoints")

# UPS
$cmbUPSCount = $mainWin.FindName("cmbUPSCount")
$stkUPS = $mainWin.FindName("stkUPS")

# CCTV
$cmbCCTVCount = $mainWin.FindName("cmbCCTVCount")
$stkCCTV = $mainWin.FindName("stkCCTV")

# Printer
$cmbPrinterCount = $mainWin.FindName("cmbPrinterCount")
$stkPrinter = $mainWin.FindName("stkPrinter")

# Buttons and Controls
$btnAddSite = $mainWin.FindName("btnAddSite")
$btnClearForm = $mainWin.FindName("btnClearForm")
$btnEditSite = $mainWin.FindName("btnEditSite")
$btnDeleteSite = $mainWin.FindName("btnDeleteSite")
$dgSites = $mainWin.FindName("dgSites")
$txtSearchSites = $mainWin.FindName("txtSearchSites")
$btnClearSearchSites = $mainWin.FindName("btnClearSearchSites")
$txtBlkSiteStatus = $mainWin.FindName("txtBlkSiteStatus")

# Lookup Controls
$txtSiteLookup = $mainWin.FindName("txtSiteLookup")
$btnLookupSite = $mainWin.FindName("btnLookupSite")
$grpSiteLookupResults = $mainWin.FindName("grpSiteLookupResults")
$stkSiteDetails = $mainWin.FindName("stkSiteDetails")

# Import/Export Controls
$txtSiteCsvFilePath = $mainWin.FindName("txtSiteCsvFilePath")
$btnBrowseSiteCsv = $mainWin.FindName("btnBrowseSiteCsv")
$btnImportSiteCsv = $mainWin.FindName("btnImportSiteCsv")
$btnExportSiteCsv = $mainWin.FindName("btnExportSiteCsv")
$txtBlkSiteImportStatus = $mainWin.FindName("txtBlkSiteImportStatus")
$pnlSiteImportProgress = $mainWin.FindName("pnlSiteImportProgress")
$pbSiteImportProgress = $mainWin.FindName("pbSiteImportProgress")
$txtSiteProgressStatus = $mainWin.FindName("txtSiteProgressStatus")
$txtSiteProgressDetails = $mainWin.FindName("txtSiteProgressDetails")

# Tab Controls
$MainTabControl = $mainWin.FindName("MainTabControl")
$SiteManagementTabControl = $mainWin.FindName("SiteManagementTabControl")

# Status Bar
$SiteStatusBar = $mainWin.FindName("SiteStatusBar")
$txtStatusBarSites = $mainWin.FindName("txtStatusBarSites")
$txtStatusBarSiteSelected = $mainWin.FindName("txtStatusBarSiteSelected")

# Search debouncing timer
$script:SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$script:SearchTimer.Add_Tick({
    Update-DataGridWithSearch
    $script:SearchTimer.Stop()
})

# ===================================================================
# IP NETWORK IDENTIFIER CONTROL REFERENCES
# ===================================================================

# IP Network Identifier Controls - with null checks
$txtIpSubnet = $mainWin.FindName("txtIpSubnet")
$txtVlanId = $mainWin.FindName("txtVlanId")
$txtVlanName = $mainWin.FindName("txtVlanName")
$txtSiteName = $mainWin.FindName("txtSiteName")
$btnAddEntry = $mainWin.FindName("btnAddEntry")
$txtIpLookup = $mainWin.FindName("txtIpLookup")
$btnLookup = $mainWin.FindName("btnLookup")
$txtBlkMatchedSubnet = $mainWin.FindName("txtBlkMatchedSubnet")
$txtBlkVlanId = $mainWin.FindName("txtBlkVlanId")
$txtBlkVlanName = $mainWin.FindName("txtBlkVlanName")
$txtBlkSiteName = $mainWin.FindName("txtBlkSiteName")
$grpLookupResults = $mainWin.FindName("grpLookupResults")
$dgSubnets = $mainWin.FindName("dgSubnets")
$btnDeleteEntry = $mainWin.FindName("btnDeleteEntry")
$txtCsvFilePath = $mainWin.FindName("txtCsvFilePath")
$btnBrowseCsv = $mainWin.FindName("btnBrowseCsv")
$btnImportCsv = $mainWin.FindName("btnImportCsv")
$btnExportCsv = $mainWin.FindName("btnExportCsv")
$txtBlkImportStatus = $mainWin.FindName("txtBlkImportStatus")
$txtBlkSearchedIp = $mainWin.FindName("txtBlkSearchedIp")
$txtSearch = $mainWin.FindName("txtSearch")
$btnClearSearch = $mainWin.FindName("btnClearSearch")
$txtStatusBarSubnets = $mainWin.FindName("txtStatusBarSubnets")
$txtStatusBarSelected = $mainWin.FindName("txtStatusBarSelected")
$pbImportProgress = $mainWin.FindName("pbImportProgress")
$txtProgressStatus = $mainWin.FindName("txtProgressStatus")
$txtProgressDetails = $mainWin.FindName("txtProgressDetails")
$pnlImportProgress = $mainWin.FindName("pnlImportProgress")
$MainStatusBar = $mainWin.FindName("MainStatusBar")

# Initialize the global device panel manager (will be created after UI loads)
$script:DeviceManager = $null

# Initialize the global field mapping manager (will be created after UI loads)
$script:FieldManager = $null

# ===================================================================
# EVENT HANDLERS SETUP
# ===================================================================

# Switch count selection changed - using centralized manager
$cmbSwitchCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'Switch' $cmbSwitchCount })
$cmbAPCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'AccessPoint' $cmbAPCount })
$cmbUPSCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'UPS' $cmbUPSCount })
$cmbCCTVCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'CCTV' $cmbCCTVCount })
$cmbPrinterCount.Add_SelectionChanged({ Handle-DeviceCountChanged 'Printer' $cmbPrinterCount })

# Backup circuit checkbox
$chkHasBackupCircuit.Add_Checked({
    if ($grdBackupCircuit) {
        $grdBackupCircuit.Visibility = "Visible"
    }
})

$chkHasBackupCircuit.Add_Unchecked({
    if ($grdBackupCircuit) {
        $grdBackupCircuit.Visibility = "Collapsed"
    }
})

# Primary modem checkbox
$chkPrimaryHasModem.Add_Checked({
    if ($stkPrimaryModem) {
        $stkPrimaryModem.Visibility = "Visible"
    }
})

$chkPrimaryHasModem.Add_Unchecked({
    if ($stkPrimaryModem) {
        $stkPrimaryModem.Visibility = "Collapsed"
    }
})

# Backup modem checkbox
$chkBackupHasModem.Add_Checked({
    if ($stkBackupModem) {
        $stkBackupModem.Visibility = "Visible"
    }
})

$chkBackupHasModem.Add_Unchecked({
    if ($stkBackupModem) {
        $stkBackupModem.Visibility = "Collapsed"
    }
})

# Site Subnet auto-population using centralized function
$txtSiteSubnet.Add_TextChanged({
    $vlanControls = @{
        VLAN100 = $txtVlan100
        VLAN101 = $txtVlan101
        VLAN102 = $txtVlan102
        VLAN103 = $txtVlan103
        VLAN104 = $txtVlan104
        VLAN105 = $txtVlan105
        VLAN106 = $txtVlan106
        VLAN107 = $txtVlan107
        VLAN108 = $txtVlan108
        VLAN109 = $txtVlan109
        VLAN110 = $txtVlan110
    }
    Update-VLANsAndIPsFromSubnet -SubnetInput $txtSiteSubnet.Text -VLANControls $vlanControls -DeviceManager $script:DeviceManager -FirewallIPControl $txtFirewallIP -SiteSubnetCodeControl $txtSiteSubnetCode
})

# Site Code auto-population using centralized function
$txtSiteCode.Add_TextChanged({
    Update-DeviceNamesFromSiteCode -SiteCode $txtSiteCode.Text -DeviceManager $script:DeviceManager -FirewallNameControl $txtFirewallName
})

# Primary circuit type selection changed
$cmbPrimaryCircuitType.Add_SelectionChanged({
    $stkPrimaryGPONElement = $mainWin.FindName("stkPrimaryGPON")
    if ($stkPrimaryGPONElement) {
        if ($cmbPrimaryCircuitType.SelectedItem -and $cmbPrimaryCircuitType.SelectedItem.Content -eq "GPON Fiber") {
            $stkPrimaryGPONElement.Visibility = "Visible"
        } else {
            $stkPrimaryGPONElement.Visibility = "Collapsed"
        }
    }
})

# Backup circuit type selection changed
$cmbBackupCircuitType.Add_SelectionChanged({
    $stkBackupGPONElement = $mainWin.FindName("stkBackupGPON")
    if ($stkBackupGPONElement) {
        if ($cmbBackupCircuitType.SelectedItem -and $cmbBackupCircuitType.SelectedItem.Content -eq "GPON Fiber") {
            $stkBackupGPONElement.Visibility = "Visible"
        } else {
            $stkBackupGPONElement.Visibility = "Collapsed"
        }
    }
})

$btnAddSite.Add_Click({
    $null = Add-Site
})

# Clear form button
$btnClearForm.Add_Click({
    Clear-SiteForm
    $txtBlkSiteStatus.Text = "Form cleared"
    $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Blue
})

# Edit site button - Now fully functional with popup window
$btnEditSite.Add_Click({
    $selectedItems = @($dgSites.SelectedItems)
    if ($selectedItems.Count -eq 1) {
        # Get the selected site data
        $selectedSite = $selectedItems[0]
        
        # Find the full site entry from the data store
        $allSites = $siteDataStore.GetAllEntries()
        $siteToEdit = $allSites | Where-Object { $_.ID -eq $selectedSite.ID }
        
        if ($siteToEdit) {
            # Show the edit window
            $editResult = Show-EditSiteWindow -SiteToEdit $siteToEdit
            
            if ($editResult) {
                # Refresh the data grid to show changes
                Update-DataGridWithSearch
            }
        } else {
            Show-ValidationError "Selected site not found in database." "Site Not Found"
        }
    } elseif ($selectedItems.Count -eq 0) {
        Show-ValidationError "Please select a site to edit." "Selection Required"
    } else {
        Show-ValidationError "Please select only one site to edit at a time." "Multiple Selection"
    }
})

# Delete site button
$btnDeleteSite.Add_Click({
    $selectedItems = @($dgSites.SelectedItems)
    if ($selectedItems.Count -gt 0) {
        $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected sites?" "Confirm Deletion" "YesNo" "Warning"
       
        if ($confirm -eq "Yes") {
            $idsToDelete = @()
            foreach ($item in $selectedItems) {
                $idsToDelete += $item.ID
            }
            if ($siteDataStore.DeleteEntries($idsToDelete)) {
                Update-DataGridWithSearch
                Show-ValidationError "Successfully deleted $($selectedItems.Count) sites." "Success"
            } else {
                Show-ValidationError "Error deleting sites." "Delete Error"
            }
        }
    } else {
        Show-ValidationError "Please select one or more sites to delete." "Selection Required"
    }
})

# Enable Delete key to remove selected sites
$dgSites.Add_PreviewKeyDown({
    param($sender, $e)
   
    # Check if Delete key was pressed
    if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
        $selectedItems = @($dgSites.SelectedItems)
       
        if ($selectedItems.Count -gt 0) {
            # Trigger delete functionality directly
            $confirm = Show-CustomDialog "Are you sure you want to delete $($selectedItems.Count) selected sites?" "Confirm Deletion" "YesNo" "Warning"
           
            if ($confirm -eq "Yes") {
                $idsToDelete = @()
                foreach ($item in $selectedItems) {
                    $idsToDelete += $item.ID
                }
                if ($siteDataStore.DeleteEntries($idsToDelete)) {
                    Update-DataGridWithSearch
                    Show-ValidationError "Successfully deleted $($selectedItems.Count) sites." "Success"
                } else {
                    Show-ValidationError "Error deleting sites." "Delete Error"
                }
            }
           
            # Mark the event as handled to prevent default behavior
            $e.Handled = $true
        }
    }
})

# Lookup site button
$btnLookupSite.Add_Click({
    $searchTerm = $txtSiteLookup.Text.Trim()
    Lookup-Site -SearchTerm $searchTerm
})

# Search functionality
$txtSearchSites.Add_TextChanged({
    $script:SearchTimer.Stop()
    $script:SearchTimer.Start()
})

$btnClearSearchSites.Add_Click({
    $txtSearchSites.Text = ""
    Update-DataGridWithSearch
})

# DataGrid selection changed
$dgSites.Add_SelectionChanged({
    $selectedItems = @($dgSites.SelectedItems)
    if ($selectedItems.Count -gt 0) {
        if ($selectedItems.Count -eq 1) {
            $txtStatusBarSiteSelected.Text = "Selected: $($selectedItems[0].SiteCode) - $($selectedItems[0].SiteName)"
        } else {
            $txtStatusBarSiteSelected.Text = "Selected: $($selectedItems.Count) sites"
        }
    } else {
        $txtStatusBarSiteSelected.Text = "Selected: None"
    }
})

# Enhanced double-click handler with better user feedback

$dgSites.Add_MouseDoubleClick({
    param($sender, $e)
    
    try {
        # Check if we actually clicked on a row (not empty space)
        $clickedItem = $dgSites.SelectedItem
        
        if ($clickedItem) {
            # Update status to show we're opening edit window
            $txtBlkSiteStatus.Text = "Opening edit window for site: $($clickedItem.SiteCode)..."
            $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Blue
            
            # Find the full site entry from the data store
            $allSites = $siteDataStore.GetAllEntries()
            $siteToEdit = $allSites | Where-Object { $_.ID -eq $clickedItem.ID }
            
            if ($siteToEdit) {
                # Show the edit window
                $editResult = Show-EditSiteWindow -SiteToEdit $siteToEdit
                
                if ($editResult) {
                    # Refresh the data grid to show changes
                    Update-DataGridWithSearch
                    $txtBlkSiteStatus.Text = "Site '$($siteToEdit.SiteCode)' updated successfully!"
                    $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Green
                } else {
                    $txtBlkSiteStatus.Text = "Edit cancelled for site: $($siteToEdit.SiteCode)"
                    $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                }
            } else {
                $txtBlkSiteStatus.Text = "Error: Selected site not found in database"
                $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Red
            }
        } else {
            # Double-clicked on empty space - provide helpful message
            $txtBlkSiteStatus.Text = "Double-click on a site row to edit it"
            $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Gray
        }
        
    } catch {
        $txtBlkSiteStatus.Text = "Error opening edit window: $_"
        $txtBlkSiteStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
})

# Tab control selection changed
$MainTabControl.Add_SelectionChanged({
    $selectedTab = $MainTabControl.SelectedItem
    
    if ($selectedTab -ne $null) {
        # Always clear status messages when switching tabs
        $txtBlkSiteStatus.Text = ""
        $txtBlkSiteImportStatus.Text = ""
        
        if ($selectedTab.Header -eq "Manage Sites") {
            $SiteStatusBar.Visibility = [System.Windows.Visibility]::Visible
            Update-DataGridWithSearch
        } else {
            $SiteStatusBar.Visibility = [System.Windows.Visibility]::Collapsed
        }
        
        # Clear Lookup Site tab when leaving it
        if ($selectedTab.Header -ne "Lookup Site") {
            $txtSiteLookup.Text = ""
            $grpSiteLookupResults.Visibility = "Collapsed"
        }
        
        # Clear Import/Export tab when leaving it
        if ($selectedTab.Header -ne "Import/Export") {
            $txtSiteCsvFilePath.Text = ""
        }
    }
})

# Enter key support for lookup
$txtSiteLookup.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        $searchTerm = $txtSiteLookup.Text.Trim()
        Lookup-Site -SearchTerm $searchTerm
    }
})

# Browse for Excel file
$btnBrowseSiteCsv.Add_Click({
    try {
        $openDialog = New-Object Microsoft.Win32.OpenFileDialog
        $openDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
        $openDialog.DefaultExt = "xlsx"
        
        if ($openDialog.ShowDialog() -eq $true) {
            $txtSiteCsvFilePath.Text = $openDialog.FileName
            $txtBlkSiteImportStatus.Text = "File selected: $($openDialog.FileName)"
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Blue
        }
    } catch {
        $txtBlkSiteImportStatus.Text = "Error selecting file: $_"
        $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
})

# Import from Excel
# Import from Excel - FIXED VERSION
$btnImportSiteCsv.Add_Click({
    try {
        $filePath = $txtSiteCsvFilePath.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            Show-CustomDialog "Please select an Excel file first." "No File Selected" "OK" "Warning"
            return
        }
        
        # IMMEDIATE FEEDBACK - Show progress panel right away
        $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Visible
        $pbSiteImportProgress.Value = 0
        $txtSiteProgressStatus.Text = "Initializing Excel application..."
        $txtSiteProgressDetails.Text = "Please wait while Excel is starting up..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $result = Import-SitesFromExcel -ExcelFilePath $filePath
        
        # Show DETAILED results in main window status area
        $txtBlkSiteImportStatus.Text = $result
        $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
        
        # Extract counts from the main window result text to ensure consistency
        $importLines = $result -split "`n"
        
        # Parse using the correct patterns from debug output
        $totalLine = $importLines | Where-Object { $_ -like "Total sites processed: *" } | Select-Object -First 1
        $updatedLine = $importLines | Where-Object { $_ -like " Updated existing: * sites" -or $_ -like " Updated: * sites" } | Select-Object -First 1  
        $noChangesLine = $importLines | Where-Object { $_ -like " No changes needed: * sites" -or $_ -like " No changes: * sites" } | Select-Object -First 1
        $newLine = $importLines | Where-Object { $_ -like " Successfully imported: * sites" } | Select-Object -First 1
        $subnetWarningsLine = $importLines | Where-Object { $_ -like " Subnet warnings: * sites" } | Select-Object -First 1
        
        # Build popup using the same text as main window
        $popupBody = ""
        if ($totalLine) { $popupBody += $totalLine.Trim() }
        if ($newLine) { $popupBody += "`n" + $newLine.Trim() }
        if ($updatedLine) { $popupBody += "`n" + $updatedLine.Trim() }
        if ($noChangesLine) { $popupBody += "`n" + $noChangesLine.Trim() }
        if ($subnetWarningsLine) { $popupBody += "`n" + $subnetWarningsLine.Trim() }
        
        # ADD VALIDATION ERROR COUNT TO EXISTING POPUP
        $errorLines = $importLines | Where-Object { $_ -like "❌*" }
        if ($errorLines.Count -gt 0) {
            $popupBody += "`nValidation errors: $($errorLines.Count) sites"
        }
        
        # Show popup with summary statistics
        Show-CustomDialog $popupBody "Import completed successfully!" "OK" "Information"
        
        # FORCE REFRESH THE DATA GRID - Multiple approaches to ensure it works
        try {
            # Method 1: Force reload data store (in case it wasn't updated properly)
            $siteDataStore.LoadData()
            
            # Method 2: Clear and reset ItemsSource
            $dgSites.ItemsSource = $null
            [System.Windows.Forms.Application]::DoEvents()  # Allow UI to process
            
            # Method 3: Use the existing update function which handles search/filter
            Update-DataGridWithSearch
            
            # Method 4: If still not working, force a complete refresh
            $allData = $siteDataStore.GetAllEntries()
            $dgSites.ItemsSource = @($allData | Sort-Object -Property @{Expression={[int]$_.ID}; Ascending=$true})
            
            # Method 5: Force UI update
            $dgSites.Items.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
                        
        } catch {
            # Fallback: try a simple refresh
            $dgSites.Items.Refresh()
        }
        
    } catch {
        $errorMsg = "Import failed: $_"
        $txtBlkSiteImportStatus.Text = $errorMsg
        $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-CustomDialog $errorMsg "Import Error" "OK" "Error"
    } finally {
        # Hide progress panel
        if ($pnlSiteImportProgress) {
            $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }
})

# Export to CSV
$btnExportSiteCsv.Add_Click({
    try {
        # Show save dialog
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        $saveDialog.DefaultExt = "xlsx"
        $saveDialog.FileName = "sites_export_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
        
        if ($saveDialog.ShowDialog() -eq $true) {
            $result = Export-SitesToExcel -FilePath $saveDialog.FileName
            
            # Get file info with smart size formatting
            $fileInfo = Get-Item $saveDialog.FileName
            $fileSizeBytes = $fileInfo.Length
            
            # Smart file size formatting
            if ($fileSizeBytes -lt 1KB) {
                $fileSize = "$fileSizeBytes bytes"
            } elseif ($fileSizeBytes -lt 1MB) {
                $fileSizeKB = [math]::Round($fileSizeBytes / 1KB, 1)
                $fileSize = "$fileSizeKB KB"
            } else {
                $fileSizeMB = [math]::Round($fileSizeBytes / 1MB, 1)
                $fileSize = "$fileSizeMB MB"
            }
            
            # Get all sites and their codes
            $allSites = $siteDataStore.GetAllEntries()
            $totalSites = $allSites.Count
            $siteCodesList = ($allSites.SiteCode | Sort-Object) -join ", "
            
            # Build main window result - EXACT format you want
            $mainResult = @"
Excel export completed successfully!
==========================================

Total sites exported: $totalSites

Site Exported :
$siteCodesList

EXPORT DETAILS:
File size: $fileSize
Location: $($saveDialog.FileName)
"@

            # Build popup result - EXACT format you want  
            $popupResult = @"
Total sites exported: $totalSites
File size: $fileSize
Location: $([System.IO.Path]::GetFileName($saveDialog.FileName))
"@

            # Show results
            $txtBlkSiteImportStatus.Text = $mainResult
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Green
            
            Show-CustomDialog $popupResult "Export completed successfully!" "OK" "Information"
            
        } else {
            $txtBlkSiteImportStatus.Text = "Export cancelled"
            $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Orange
        }
        
    } catch {
        $errorMsg = "Export failed: $_"
        $txtBlkSiteImportStatus.Text = $errorMsg
        $txtBlkSiteImportStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-CustomDialog $errorMsg "Export Error" "OK" "Error"
    }
})


# Import edit window functions
try {
    $editWindowPath = Join-Path $scriptPath "EditSiteWindow.ps1"    
    if (Test-Path $editWindowPath) {
        . $editWindowPath
    } 
    else {
        $errorMsg = "EditSiteWindow.ps1 not found at: $editWindowPath"
        Show-MessageBox $errorMsg "Module Error" "OK" "Error"
        exit 1
    }
}
catch {
    $errorMsg = "Failed to load EditSiteWindow.ps1: $_"
    Show-MessageBox $errorMsg "Module Error" "OK" "Error"
    exit 1
}

# ===================================================================
# APPLICATION INITIALIZATION
# ===================================================================

# Initialize the data store first
$siteDataStore = [SiteDataStore]::new()

# Initialize the IP subnet data store
$subnetDataStore = [SubnetDataStore]::new()

# Initialize managers to null first
$script:DeviceManager = $null
$script:FieldManager = $null

# Add window loaded event to initialize managers after UI is ready
$mainWin.Add_Loaded({
    try {
        # Initialize managers after window is fully loaded
        $script:DeviceManager = [DevicePanelManager]::new($mainWin)
        $script:FieldManager = [FieldMappingManager]::new($mainWin)
        
        # Set initial visibility states
        if ($grpSiteLookupResults) { $grpSiteLookupResults.Visibility = "Collapsed" }
        if ($grdBackupCircuit) { $grdBackupCircuit.Visibility = "Collapsed" }
        if ($stkPrimaryModem) { $stkPrimaryModem.Visibility = "Collapsed" }
        if ($stkBackupModem) { $stkBackupModem.Visibility = "Collapsed" }
        if ($pnlSiteImportProgress) { $pnlSiteImportProgress.Visibility = "Collapsed" }
        if ($SiteStatusBar) { $SiteStatusBar.Visibility = "Collapsed" }
        
        # Initialize IP Network components
        if ($grpLookupResults) { $grpLookupResults.Visibility = "Collapsed" }
        if ($pnlImportProgress) { $pnlImportProgress.Visibility = "Collapsed" }
        if ($MainStatusBar) { $MainStatusBar.Visibility = "Collapsed" }
        
        # Initialize the data grids
        Update-DataGridWithSearch
        
        # Initialize IP subnet data grid ONLY if controls exist
        if ($dgSubnets -ne $null) {
            Update-SubnetDataGridWithSearch
        }
        
        # Initialize IP Network event handlers ONLY if controls exist
        if ($btnAddEntry -ne $null) {
            Initialize-IPNetworkEventHandlers
        } else {
        }
        
        # Initialize phone formatting
        $txtMainContactPhone.Add_LostFocus({ $this.Text = Format-PhoneNumber $this.Text })
        $txtSecondContactPhone.Add_LostFocus({ $this.Text = Format-PhoneNumber $this.Text })
        
    } catch {
        Show-MessageBox "Error initializing application: $_" "Initialization Error" "OK" "Error"
    }
})

# Show the window
    try {
        $mainWin.ShowDialog() | Out-Null
    } catch {
}