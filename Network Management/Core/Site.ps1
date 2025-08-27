# Load Assemblies
Add-Type -AssemblyName WindowsBase, PresentationFramework, PresentationCore, System.Windows.Forms

# Define XAML file path
$xamlFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) ".." | Join-Path -ChildPath "UI" | Join-Path -ChildPath "NetworkManagement.xaml"

# Import the data models first
try {
    # Get the script directory correctly
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    # Construct the path to DataModels.ps1
    $dataModelsPath = Join-Path $scriptPath "DataModels.ps1"
        
    if (Test-Path $dataModelsPath) {
        # Dot-source the data models
        . $dataModelsPath
    } 
    else {
        $errorMsg = "DataModels.ps1 not found at: $dataModelsPath"
        [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
        exit 1
    }
}
catch {
    $errorMsg = "Failed to load DataModels.ps1: $_`n`nFull path tried: $dataModelsPath"
    [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
    exit 1
}

# Import the Import/Export module
try {
    # Construct the path to SiteImportExport.ps1
    $importExportPath = Join-Path $scriptPath "SiteImportExport.ps1"
        
    if (Test-Path $importExportPath) {
        # Dot-source the import/export module
        . $importExportPath
    } 
    else {
        $errorMsg = "SiteImportExport.ps1 not found at: $importExportPath"
        [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
        exit 1
    }
}
catch {
    $errorMsg = "Failed to load SiteImportExport.ps1: $_`n`nFull path tried: $importExportPath"
    [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
    exit 1
}

# Import the IP Network module
try {
    $ipNetworkPath = Join-Path $scriptPath "IPNetworkModule.ps1"    
    if (Test-Path $ipNetworkPath) {
        . $ipNetworkPath
    } 
    else {
        $errorMsg = "IPNetworkModule.ps1 not found at: $ipNetworkPath"
        [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
        exit 1
    }
}
catch {
    $errorMsg = "Failed to load IPNetworkModule.ps1: $_"
    [System.Windows.MessageBox]::Show($errorMsg, "Module Error", "OK", "Error")
    exit 1
}

# ===================================================================
# PHONE NUMBER CONVERTER CLASS
# ===================================================================

# Phone number converter for XAML binding
class PhoneNumberConverter : System.Windows.Data.IValueConverter {
    [object] Convert([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return Format-PhoneNumber $value
    }
    
    [object] ConvertBack([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return $value
    }
}

# ===================================================================
# UTILITY FUNCTIONS
# ===================================================================

# Safely release COM objects to prevent memory leaks
function Release-ComObject {
    param($ComObject)
    if ($ComObject) {
        try {
            $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            # Keep releasing until reference count is 0
            while ($refCount -gt 0) {
                $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            }
        } catch {
            # Silent fail on release
        }
    }
    return $null
}

# Get safe string value from object, returns empty string if null
function Get-SafeValue {
    param([object]$Value)
    if ($Value) { return $Value.ToString() } else { return "" }
}

# Format phone number to standard format: +1 (xxx) xxx-xxxx
function Format-PhoneNumber {
    param([string]$PhoneNumber)
    
    if ([string]::IsNullOrWhiteSpace($PhoneNumber)) { return "" }
    
    # Remove all non-digits
    $digits = $PhoneNumber -replace '\D', ''
    
    # Format 10 digits: xxxxxxxxxx -> +1 (xxx) xxx-xxxx
    if ($digits.Length -eq 10) {
        return "+1 ($($digits.Substring(0,3))) $($digits.Substring(3,3))-$($digits.Substring(6,4))"
    }
    
    # Return original if not 10 digits
    return $PhoneNumber
}

# ===================================================================
# CENTRALIZED AUTO-POPULATION FUNCTIONS
# ===================================================================

# Centralized function to update device names from site code
function Update-DeviceNamesFromSiteCode {
    param(
        [string]$SiteCode,
        [object]$DeviceManager,
        [object]$FirewallNameControl
    )
    
    if ([string]::IsNullOrWhiteSpace($SiteCode)) { return }
    
    # Update all device types using the device manager
    $DeviceManager.UpdateDeviceNamesFromSiteCode('Switch', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('AccessPoint', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('UPS', $SiteCode)
    $DeviceManager.UpdateDeviceNamesFromSiteCode('CCTV', $SiteCode)
    
    # Update firewall name (not managed by DeviceManager)
    if ($FirewallNameControl -and -not [string]::IsNullOrWhiteSpace($SiteCode)) {
        $siteCodeUpper = $SiteCode.Trim().ToUpper()
        $FirewallNameControl.Text = "$siteCodeUpper-FW"
    }
}

# Centralized function to update VLANs and IPs from subnet
function Update-VLANsAndIPsFromSubnet {
    param(
        [string]$SubnetInput,
        [hashtable]$VLANControls,
        [object]$DeviceManager,
        [object]$FirewallIPControl,
        [object]$SiteSubnetCodeControl
    )
    
    if ([string]::IsNullOrWhiteSpace($SubnetInput)) { return }
    
    # Parse subnet (e.g., "10.107.0.0" -> "10.107")
    if ($SubnetInput -match '^(\d+\.\d+)\.') {
        $baseSubnet = $matches[1]
        
        # Auto-populate VLAN fields
        if ($VLANControls.VLAN100) { $VLANControls.VLAN100.Text = "$baseSubnet.10.0" }
        if ($VLANControls.VLAN101) { $VLANControls.VLAN101.Text = "$baseSubnet.20.0" }
        if ($VLANControls.VLAN102) { $VLANControls.VLAN102.Text = "$baseSubnet.102.0" }
        if ($VLANControls.VLAN103) { $VLANControls.VLAN103.Text = "$baseSubnet.103.0" }
        if ($VLANControls.VLAN104) { $VLANControls.VLAN104.Text = "$baseSubnet.40.0" }
        if ($VLANControls.VLAN105) { $VLANControls.VLAN105.Text = "$baseSubnet.50.0" }
        if ($VLANControls.VLAN106) { $VLANControls.VLAN106.Text = "$baseSubnet.60.0" }
        if ($VLANControls.VLAN107) { $VLANControls.VLAN107.Text = "$baseSubnet.70.0" }
        if ($VLANControls.VLAN108) { $VLANControls.VLAN108.Text = "$baseSubnet.80.0" }
        if ($VLANControls.VLAN109) { $VLANControls.VLAN109.Text = "$baseSubnet.90.0" }
        if ($VLANControls.VLAN110) { $VLANControls.VLAN110.Text = "$baseSubnet.110.0" }
        
        # Auto-fill firewall IP
        if ($FirewallIPControl) {
            $firewallIP = "$baseSubnet.20.1"
            $FirewallIPControl.Text = $firewallIP
        }
        
        # Auto-fill device IPs using device manager
        $DeviceManager.UpdateDeviceIPsFromSubnet('Switch', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('AccessPoint', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('UPS', $baseSubnet)
        $DeviceManager.UpdateDeviceIPsFromSubnet('CCTV', $baseSubnet)
    }
    
    # Auto-populate site subnet code
    if ($SiteSubnetCodeControl -and $SubnetInput -match '^(\d+)\.(\d+)\.') {
        $secondOctet = $matches[2]
        $siteSubnetCode = [int]$secondOctet
        
        # Only auto-fill if the field is empty
        if ([string]::IsNullOrWhiteSpace($SiteSubnetCodeControl.Text)) {
            $SiteSubnetCodeControl.Text = $siteSubnetCode.ToString()
        }
    }
}

# ===================================================================
# CENTRALIZED SITE VALIDATION FUNCTIONS
# ===================================================================
# Centralized function to validate basic site information
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
function Set-ComboBoxValue {
    param(
        [System.Windows.Controls.ComboBox]$ComboBox,
        [object]$Value,  # Can be string, int, or any type
        [switch]$ByContent = $false  # If true, match by Content property, otherwise by value
    )
    
    if ($ComboBox -eq $null) { return }
    
    if ($Value -eq $null -or $Value -eq "") {
        $ComboBox.SelectedIndex = -1
        return
    }
    
    # Convert value to string for comparison
    $searchValue = $Value.ToString().Trim()
    
    for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
        $itemValue = ""
        
        if ($ByContent -and $ComboBox.Items[$i].Content) {
            $itemValue = $ComboBox.Items[$i].Content.ToString().Trim()
        } elseif ($ComboBox.Items[$i]) {
            $itemValue = $ComboBox.Items[$i].ToString().Trim()
        }
        
        # Try exact match first, then try numeric comparison for numbers
        if ($itemValue -eq $searchValue) {
            $ComboBox.SelectedIndex = $i
            return
        }
        
        # Try numeric comparison if both values can be converted to numbers
        try {
            $numericSearch = [decimal]$searchValue
            $numericItem = [decimal]$itemValue
            if ($numericSearch -eq $numericItem) {
                $ComboBox.SelectedIndex = $i
                return
            }
        } catch {
            # Not numeric, continue with string comparison
        }
    }
    
    # No match found
    $ComboBox.SelectedIndex = -1
}

# ===================================================================
# VALIDATION UTILITY CLASS
# ===================================================================

# Validation utility class to consolidate validation logic
class ValidationUtility {
    static [bool] ValidateIP([string]$IPAddress) {
        if ([string]::IsNullOrWhiteSpace($IPAddress)) { return $true }
        try { 
            $null = [System.Net.IPAddress]::Parse($IPAddress.Trim())
            return $true 
        } catch { 
            return $false 
        }
    }
    
    static [void] ValidateDeviceIPs([SiteEntry]$Site) {
        foreach ($switch in $Site.Switches) {
            if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($switch.ManagementIP)) {
                    throw "Invalid Switch IP: $($switch.ManagementIP)"
                }
            }
        }
        
        foreach ($ap in $Site.AccessPoints) {
            if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($ap.ManagementIP)) {
                    throw "Invalid Access Point IP: $($ap.ManagementIP)"
                }
            }
        }
        
        foreach ($ups in $Site.UPSDevices) {
            if (-not [string]::IsNullOrWhiteSpace($ups.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($ups.ManagementIP)) {
                    throw "Invalid UPS IP: $($ups.ManagementIP)"
                }
            }
        }
        
        foreach ($cctv in $Site.CCTVDevices) {
            if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP)) {
                if (-not [ValidationUtility]::ValidateIP($cctv.ManagementIP)) {
                    throw "Invalid CCTV IP: $($cctv.ManagementIP)"
                }
            }
        }
        foreach ($printer in $Site.PrinterDevices) {
        if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP)) {
            if (-not [ValidationUtility]::ValidateIP($printer.ManagementIP)) {
                throw "Invalid Printer IP: $($printer.ManagementIP)"
                }
            }
        }
    }
}

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

# ===================================================================
# DEVICE MANAGEMENT CLASSES
# ===================================================================

# Configuration class for different device types (switches, APs, UPS, CCTV)
class DeviceConfiguration {
    [string]$Type
    [string]$Prefix
    [string]$VLANSubnet
    [int]$IPStartOffset
    [int]$MaxCount
    [string[]]$Fields
    [hashtable]$FieldLabels
    [string]$HeaderTemplate
    
    DeviceConfiguration([string]$type, [string]$prefix, [string]$vlanSubnet, [int]$ipOffset, [int]$maxCount, [string[]]$fields, [hashtable]$labels, [string]$headerTemplate) {
        $this.Type = $type
        $this.Prefix = $prefix
        $this.VLANSubnet = $vlanSubnet
        $this.IPStartOffset = $ipOffset
        $this.MaxCount = $maxCount
        $this.Fields = $fields
        $this.FieldLabels = $labels
        $this.HeaderTemplate = $headerTemplate
    }
}

# ===================================================================
# UNIVERSAL DEVICE PANEL FACTORY
# ===================================================================

# Universal factory for creating device panels - eliminates duplication
class UniversalDevicePanelFactory {
    static [System.Windows.Controls.GroupBox] CreateDevicePanel([DeviceConfiguration]$Config, [int]$DeviceNumber, [string]$ControlPrefix = "") {
        try {
            $groupBox = New-Object System.Windows.Controls.GroupBox
            $groupBox.Header = $Config.HeaderTemplate -f $DeviceNumber
            $groupBox.Margin = New-Object System.Windows.Thickness(0,0,0,10)
            
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = New-Object System.Windows.Thickness(5)
            
            # Create 2 columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(1, 'Auto')
            $col2 = New-Object System.Windows.Controls.ColumnDefinition  
            $col2.Width = New-Object System.Windows.GridLength(1, 'Star')
            $grid.ColumnDefinitions.Add($col1)
            $grid.ColumnDefinitions.Add($col2)
            
            # Create rows for fields
            for ($i = 0; $i -lt $Config.Fields.Count; $i++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = New-Object System.Windows.GridLength(1, 'Auto')
                $grid.RowDefinitions.Add($row)
            }
            
            # Add fields dynamically
            for ($i = 0; $i -lt $Config.Fields.Count; $i++) {
                $field = $Config.Fields[$i]
                $label = $Config.FieldLabels[$field]
                
                # Create label
                $lblControl = New-Object System.Windows.Controls.Label
                $lblControl.Content = $label
                [System.Windows.Controls.Grid]::SetRow($lblControl, $i)
                [System.Windows.Controls.Grid]::SetColumn($lblControl, 0)
                $grid.Children.Add($lblControl) | Out-Null
                
                # Create textbox with configurable prefix
                $txtControl = New-Object System.Windows.Controls.TextBox
                $txtControl.Name = "$ControlPrefix$($Config.Type)$DeviceNumber$field"
                $txtControl.Margin = New-Object System.Windows.Thickness(0,2,0,2)
                [System.Windows.Controls.Grid]::SetRow($txtControl, $i)
                [System.Windows.Controls.Grid]::SetColumn($txtControl, 1)
                $grid.Children.Add($txtControl) | Out-Null
            }
            
            $groupBox.Content = $grid
            return $groupBox
            
        } catch {
            [System.Windows.MessageBox]::Show("Error creating $($Config.Type) panel: $_", "Panel Creation Error", "OK", "Error")
            return $null
        }
    }
}

# ===================================================================
# UNIVERSAL DATA COLLECTOR
# ===================================================================

# Universal data collector - eliminates GetDeviceDataFromUI duplication
class UniversalDataCollector {
    static [object] CollectDeviceData([DeviceConfiguration]$Config, [object]$StackPanel, [object]$ComboBox, [string]$ControlPrefix = "txt") {
        # Determine device type class name
        $className = switch ($Config.Type) {
            'Switch' { 'SwitchInfo' }
            'AccessPoint' { 'AccessPointInfo' }
            'UPS' { 'UPSInfo' }
            'CCTV' { 'CCTVInfo' }
            'Printer' { 'PrinterInfo' }
        }
        
        $devices = New-Object "System.Collections.Generic.List[$className]"
        Write-Host "DEBUG: Creating list for className: '$className', DeviceType: '$($Config.Type)'"
        $deviceCount = if ($ComboBox.SelectedItem) { [int]$ComboBox.SelectedItem.Content } else { 0 }
        
        for ($i = 1; $i -le $deviceCount; $i++) {
            $device = New-Object $className
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $i)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$i$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control) {
                            $device.$field = $control.Text.Trim()
                        }
                    }
                    break
                }
            }
            $devices.Add($device)
        }
        return $devices
    }
    
    # Helper method to find control in panel
    static [object] FindControlInPanel([object]$GroupBox, [string]$ControlName) {
        $grid = $GroupBox.Content
        foreach ($control in $grid.Children) {
            if ($control.Name -eq $ControlName) {
                return $control
            }
        }
        return $null
    }
    
    # Universal data populator
    static [void] PopulateDevicePanels([DeviceConfiguration]$Config, [object]$StackPanel, [array]$DeviceList, [string]$ControlPrefix = "txt") {
        if (-not $DeviceList -or $DeviceList.Count -eq 0) { return }
        
        for ($i = 0; $i -lt $DeviceList.Count; $i++) {
            $deviceNum = $i + 1
            $device = $DeviceList[$i]
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $deviceNum)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$deviceNum$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control -and $device.$field) {
                            $control.Text = $device.$field
                        }
                    }
                    break
                }
            }
        }
    }
    
    # Universal data restoration
    static [void] RestoreDeviceData([DeviceConfiguration]$Config, [object]$StackPanel, [array]$ExistingData, [int]$NewCount, [string]$ControlPrefix = "txt") {
        if (-not $ExistingData) { return }
        
        $maxRestore = [Math]::Min($ExistingData.Count, $NewCount)
        
        for ($i = 0; $i -lt $maxRestore; $i++) {
            $deviceNum = $i + 1
            $deviceData = $ExistingData[$i]
            
            foreach ($groupBox in $StackPanel.Children) {
                if ($groupBox.Header -eq ($Config.HeaderTemplate -f $deviceNum)) {
                    foreach ($field in $Config.Fields) {
                        $controlName = "$ControlPrefix$($Config.Type)$deviceNum$field"
                        $control = [UniversalDataCollector]::FindControlInPanel($groupBox, $controlName)
                        if ($control -and $deviceData.$field) {
                            $control.Text = $deviceData.$field
                        }
                    }
                    break
                }
            }
        }
    }
}

# Centralized manager for device panel creation and management
class DevicePanelManager {
    [hashtable]$Configurations
    [hashtable]$StackPanels
    [hashtable]$ComboBoxes
    [object]$MainWindow
    
    DevicePanelManager([object]$mainWindow) {
        $this.MainWindow = $mainWindow
        $this.InitializeConfigurations()
        $this.InitializeUIReferences()
    }
    
    [void] InitializeConfigurations() {
        $this.Configurations = @{
            'Switch' = [DeviceConfiguration]::new(
                'Switch',
                'SWT', 
                '.20',
                5,
                10,
                @('ManagementIP', 'Name', 'AssetTag', 'Version', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'AssetTag' = 'Asset Tag:'
                    'Version' = 'Version:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Switch {0}'
            )
            'AccessPoint' = [DeviceConfiguration]::new(
                'AccessPoint',
                'AP',
                '.20',
                100,
                10,
                @('ManagementIP', 'Name', 'AssetTag', 'Version', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'AssetTag' = 'Asset Tag:'
                    'Version' = 'Version:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Access Point {0}'
            )
            'UPS' = [DeviceConfiguration]::new(
                'UPS',
                'UPS',
                '.102',
                100,
                5,
                @('ManagementIP', 'Name'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                },
                'UPS {0}'
            )
            'CCTV' = [DeviceConfiguration]::new(
                'CCTV',
                'CAM',
                '.110',
                50,
                15,
                @('ManagementIP', 'Name', 'SerialNumber'),
                @{
                    'ManagementIP' = 'Management IP:'
                    'Name' = 'Name:'
                    'SerialNumber' = 'Serial Number:'
                },
                'Camera {0}'
            )
            'Printer' = [DeviceConfiguration]::new(
            'Printer',
            'PRT',
            '.102',
            50,
            6,
            @('ManagementIP', 'Name', 'Model', 'SerialNumber'),
            @{
                'ManagementIP' = 'Management IP:'
                'Name' = 'Name:'
                'Model' = 'Model:'
                'SerialNumber' = 'Serial Number:'
            },
            'Printer {0}'
            )
        }
    }
    
    [void] InitializeUIReferences() {
        $this.StackPanels = @{
            'Switch' = $this.MainWindow.FindName("stkSwitches")
            'AccessPoint' = $this.MainWindow.FindName("stkAccessPoints") 
            'UPS' = $this.MainWindow.FindName("stkUPS")
            'CCTV' = $this.MainWindow.FindName("stkCCTV")
            'Printer' = $this.MainWindow.FindName("stkPrinter")
        }
        
        $this.ComboBoxes = @{
            'Switch' = $this.MainWindow.FindName("cmbSwitchCount")
            'AccessPoint' = $this.MainWindow.FindName("cmbAPCount")
            'UPS' = $this.MainWindow.FindName("cmbUPSCount") 
            'CCTV' = $this.MainWindow.FindName("cmbCCTVCount")
            'Printer' = $this.MainWindow.FindName("cmbPrinterCount")
        }
    }
    
        
    # Universal panel update
    [void] UpdateDevicePanels([string]$deviceType, [int]$count) {
        try {
            $stackPanel = $this.StackPanels[$deviceType]

            if (-not $stackPanel) { return }
            
            # Save existing data
            $existingData = @()
            Write-Host "DEBUG: Got existing data, calling RestoreDeviceData for $deviceType"
            
            # Clear existing panels
            $stackPanel.Children.Clear()
            $stackPanel.RowDefinitions.Clear()
            $stackPanel.ColumnDefinitions.Clear()
            
            if ($count -eq 0) { return }
            
            # Calculate layout
            $numRows = [Math]::Ceiling($count / 2)
            
            # Setup grid layout
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(1, 'Star')
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = New-Object System.Windows.GridLength(1, 'Star')
            $stackPanel.ColumnDefinitions.Add($col1)
            $stackPanel.ColumnDefinitions.Add($col2)
            
            for ($r = 0; $r -lt $numRows; $r++) {
                $row = New-Object System.Windows.Controls.RowDefinition
                $row.Height = New-Object System.Windows.GridLength(1, 'Auto')
                $stackPanel.RowDefinitions.Add($row)
            }
            
            # Create panels
            for ($i = 1; $i -le $count; $i++) {
                Write-Host "DEBUG: Starting loop iteration $i for $deviceType"
                $config = $this.Configurations[$deviceType]

                $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
                $panel = [UniversalDevicePanelFactory]::CreateDevicePanel($config, $i, $controlPrefix)
                Write-Host "DEBUG: Created panel $i for $deviceType, panel is null: $($panel -eq $null)"
                if ($panel) {
                    # Position in grid
                    $row = [Math]::Floor(($i - 1) / 2)
                    $col = ($i - 1) % 2
                    
                    [System.Windows.Controls.Grid]::SetRow($panel, $row)
                    [System.Windows.Controls.Grid]::SetColumn($panel, $col)
                    $panel.Margin = New-Object System.Windows.Thickness(0,0,10,10)
                    
                    $stackPanel.Children.Add($panel) | Out-Null
                }
            }
            
            # Use universal data restorer
            $config = $this.Configurations[$deviceType]

            $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
            [UniversalDataCollector]::RestoreDeviceData($config, $stackPanel, $existingData, $count, $controlPrefix)
            
        } catch {
            [System.Windows.MessageBox]::Show("Error updating $deviceType panels: $_", "Panel Update Error", "OK", "Error")
        }
    }
    
    # Universal data collection
    [object] GetDeviceDataFromUI([string]$deviceType) {
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $comboBox = $this.ComboBoxes[$deviceType]
        
        # Determine device type class name
        $className = switch ($deviceType) {
            'Switch' { 'SwitchInfo' }
            'AccessPoint' { 'AccessPointInfo' }
            'UPS' { 'UPSInfo' }
            'CCTV' { 'CCTVInfo' }
        }
        
        $devices = New-Object "System.Collections.Generic.List[$className]"
        Write-Host "DEBUG: Creating list for className: '$className', DeviceType: '$($Config.Type)'"
        $deviceCount = if ($comboBox.SelectedItem) { [int]$comboBox.SelectedItem.Content } else { 0 }
        
        for ($i = 1; $i -le $deviceCount; $i++) {
            $device = New-Object $className
            
            foreach ($groupBox in $stackPanel.Children) {
                if ($groupBox.Header -eq ($config.HeaderTemplate -f $i)) {
                    foreach ($field in $config.Fields) {
                        $controlName = "txt$deviceType$i$field"
                        $control = $this.FindControlInPanel($groupBox, $controlName)
                        if ($control) {
                            $device.$field = $control.Text
                        }
                    }
                    break
                }
            }
            $devices.Add($device)
        }
        return $devices
    }

    [void] RestoreDeviceData([string]$deviceType, [array]$existingData, [int]$newCount) {
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $controlPrefix = if ($this -is [EditDevicePanelManager]) { "txtEdit" } else { "txt" }
        [UniversalDataCollector]::RestoreDeviceData($config, $stackPanel, $existingData, $newCount, $controlPrefix)
    }
    
    # Helper method to find control in panel
    [object] FindControlInPanel([object]$groupBox, [string]$controlName) {
        $grid = $groupBox.Content
        foreach ($control in $grid.Children) {
            if ($control.Name -eq $controlName) {
                return $control
            }
        }
        return $null
    }
    
    
    # Universal auto-naming
    [void] UpdateDeviceNamesFromSiteCode([string]$deviceType, [string]$siteCode) {
        if ([string]::IsNullOrWhiteSpace($siteCode)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $siteCode = $siteCode.Trim().ToUpper()
        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = $matches[1]
                $paddedNumber = $deviceNumber.PadLeft(3, '0')
                $deviceName = "$siteCode-$($config.Prefix)-$paddedNumber"
                
                $nameControl = $this.FindControlInPanel($groupBox, "txt$deviceType${deviceNumber}Name")
                if ($nameControl) {
                    $nameControl.Text = $deviceName
                }
            }
        }
    }
    
    # Universal IP auto-population  
    [void] UpdateDeviceIPsFromSubnet([string]$deviceType, [string]$baseSubnet) {
        if ([string]::IsNullOrWhiteSpace($baseSubnet)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = [int]$matches[1]
                $deviceIP = "$baseSubnet$($config.VLANSubnet).$($deviceNumber + $config.IPStartOffset - 1)"
                
                $ipControl = $this.FindControlInPanel($groupBox, "txt$deviceType${deviceNumber}ManagementIP")
                if ($ipControl -and [string]::IsNullOrWhiteSpace($ipControl.Text)) {
                    $ipControl.Text = $deviceIP
                }
            }
        }
    }
}

# ===================================================================
# ADDITIONAL HELPER FUNCTIONS FOR EDIT WINDOW
# ===================================================================

# Enhanced DevicePanelManager to work with edit window naming convention
class EditDevicePanelManager : DevicePanelManager {
    EditDevicePanelManager([object]$editWindow) : base($editWindow) {
        $this.InitializeEditUIReferences()
    }
    
    [void] InitializeEditUIReferences() {
        $this.StackPanels = @{
            'Switch' = $this.MainWindow.FindName("stkEditSwitches")
            'AccessPoint' = $this.MainWindow.FindName("stkEditAccessPoints") 
            'UPS' = $this.MainWindow.FindName("stkEditUPS")
            'CCTV' = $this.MainWindow.FindName("stkEditCCTV")
            'Printer' = $this.MainWindow.FindName("stkEditPrinter")
        }
        
        $this.ComboBoxes = @{
            'Switch' = $this.MainWindow.FindName("cmbEditSwitchCount")
            'AccessPoint' = $this.MainWindow.FindName("cmbEditAPCount")
            'UPS' = $this.MainWindow.FindName("cmbEditUPSCount") 
            'CCTV' = $this.MainWindow.FindName("cmbEditCCTVCount")
            'Printer' = $this.MainWindow.FindName("cmbEditPrinterCount")
        }
    }
    
    # Override UpdateDeviceNamesFromSiteCode to use Edit naming convention
    [void] UpdateDeviceNamesFromSiteCode([string]$deviceType, [string]$siteCode) {
        if ([string]::IsNullOrWhiteSpace($siteCode)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        $siteCode = $siteCode.Trim().ToUpper()
        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = $matches[1]
                $paddedNumber = $deviceNumber.PadLeft(3, '0')
                $deviceName = "$siteCode-$($config.Prefix)-$paddedNumber"
                
                # Use Edit naming convention
                $nameControl = $this.FindControlInPanel($groupBox, "txtEdit$deviceType${deviceNumber}Name")
                if ($nameControl) {
                    $nameControl.Text = $deviceName
                }
            }
        }
    }
    
    # Override UpdateDeviceIPsFromSubnet to use Edit naming convention  
    [void] UpdateDeviceIPsFromSubnet([string]$deviceType, [string]$baseSubnet) {
        if ([string]::IsNullOrWhiteSpace($baseSubnet)) { return }
        
        $config = $this.Configurations[$deviceType]

        $stackPanel = $this.StackPanels[$deviceType]

        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -match ($config.HeaderTemplate -f '(\d+)')) {
                $deviceNumber = [int]$matches[1]
                $deviceIP = "$baseSubnet$($config.VLANSubnet).$($deviceNumber + $config.IPStartOffset - 1)"
                
                # Use Edit naming convention
                $ipControl = $this.FindControlInPanel($groupBox, "txtEdit$deviceType${deviceNumber}ManagementIP")
                if ($ipControl) {
                $ipControl.Text = $deviceIP
                }
            }
        }
    }
}

# ===================================================================
# FIELD MAPPING MANAGEMENT CLASS
# ===================================================================

# Centralized manager for form field mappings and validation
class FieldMappingManager {
    [hashtable]$MappingGroups
    [object]$MainWindow
    
    FieldMappingManager([object]$mainWindow) {
        $this.MainWindow = $mainWindow
        $this.InitializeMappingGroups()
    }
    
    [void] InitializeMappingGroups() {
        $this.MappingGroups = @{
            'BasicInfo' = @(
                @{Control = 'txtSiteCode'; Property = 'SiteCode'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteSubnet'; Property = 'SiteSubnet'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteSubnetCode'; Property = 'SiteSubnetCode'; Required = $false; Type = 'Text'},
                @{Control = 'txtSiteNameManage'; Property = 'SiteName'; Required = $true; Type = 'Text'},
                @{Control = 'txtSiteAddress'; Property = 'SiteAddress'; Required = $false; Type = 'Text'},
                @{Control = 'txtMainContactName'; Property = 'MainContactName'; Required = $false; Type = 'Text'},
                @{Control = 'txtMainContactPhone'; Property = 'MainContactPhone'; Required = $false; Type = 'Text'},
                @{Control = 'txtSecondContactName'; Property = 'SecondContactName'; Required = $false; Type = 'Text'},
                @{Control = 'txtSecondContactPhone'; Property = 'SecondContactPhone'; Required = $false; Type = 'Text'}
            )
            'Firewall' = @(
                @{Control = 'txtFirewallIP'; Property = 'FirewallIP'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtFirewallName'; Property = 'FirewallName'; Required = $false; Type = 'Text'},
                @{Control = 'txtFirewallVersion'; Property = 'FirewallVersion'; Required = $false; Type = 'Text'},
                @{Control = 'txtFirewallSN'; Property = 'FirewallSN'; Required = $false; Type = 'Text'}
            )
            'VLANs' = @(
                @{Control = 'txtVlan100'; Property = 'VLAN100_Servers'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan101'; Property = 'VLAN101_NetworkDevices'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan102'; Property = 'VLAN102_UserDevices'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan103'; Property = 'VLAN103_UserDevices2'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan104'; Property = 'VLAN104_VOIP'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan105'; Property = 'VLAN105_WiFiCorp'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan106'; Property = 'VLAN106_WiFiBYOD'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan107'; Property = 'VLAN107_WiFiGuest'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan108'; Property = 'VLAN108_Spare'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan109'; Property = 'VLAN109_DMZ'; Required = $false; Type = 'Text'},
                @{Control = 'txtVlan110'; Property = 'VLAN110_CCTV'; Required = $false; Type = 'Text'}
            )
            'PrimaryCircuit' = @(
                @{Control = 'txtPrimaryVendor'; Property = 'Vendor'; Required = $false; Type = 'Text'},
                @{Control = 'cmbPrimaryCircuitType'; Property = 'CircuitType'; Required = $false; Type = 'ComboBox'},
                @{Control = 'txtPrimaryCircuitID'; Property = 'CircuitID'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryDownloadSpeed'; Property = 'DownloadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryUploadSpeed'; Property = 'UploadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryIPAddress'; Property = 'IPAddress'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimarySubnetMask'; Property = 'SubnetMask'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryDefaultGateway'; Property = 'DefaultGateway'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryDNS1'; Property = 'DNS1'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryDNS2'; Property = 'DNS2'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtPrimaryRouterModel'; Property = 'RouterModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryRouterName'; Property = 'RouterName'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryRouterSN'; Property = 'RouterSN'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryPPPoEUsername'; Property = 'PPPoEUsername'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryPPPoEPassword'; Property = 'PPPoEPassword'; Required = $false; Type = 'Text'},
                @{Control = 'chkPrimaryHasModem'; Property = 'HasModem'; Required = $false; Type = 'CheckBox'},
                @{Control = 'txtPrimaryModemModel'; Property = 'ModemModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryModemName'; Property = 'ModemName'; Required = $false; Type = 'Text'},
                @{Control = 'txtPrimaryModemSN'; Property = 'ModemSN'; Required = $false; Type = 'Text'}
            )
            'BackupCircuit' = @(
                @{Control = 'txtBackupVendor'; Property = 'Vendor'; Required = $false; Type = 'Text'},
                @{Control = 'cmbBackupCircuitType'; Property = 'CircuitType'; Required = $false; Type = 'ComboBox'},
                @{Control = 'txtBackupCircuitID'; Property = 'CircuitID'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupDownloadSpeed'; Property = 'DownloadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupUploadSpeed'; Property = 'UploadSpeed'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupIPAddress'; Property = 'IPAddress'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupSubnetMask'; Property = 'SubnetMask'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupDefaultGateway'; Property = 'DefaultGateway'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupDNS1'; Property = 'DNS1'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupDNS2'; Property = 'DNS2'; Required = $false; Type = 'Text'; Validator = 'IP'},
                @{Control = 'txtBackupRouterModel'; Property = 'RouterModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupRouterName'; Property = 'RouterName'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupRouterSN'; Property = 'RouterSN'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupPPPoEUsername'; Property = 'PPPoEUsername'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupPPPoEPassword'; Property = 'PPPoEPassword'; Required = $false; Type = 'Text'},
                @{Control = 'chkBackupHasModem'; Property = 'HasModem'; Required = $false; Type = 'CheckBox'},
                @{Control = 'txtBackupModemModel'; Property = 'ModemModel'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupModemName'; Property = 'ModemName'; Required = $false; Type = 'Text'},
                @{Control = 'txtBackupModemSN'; Property = 'ModemSN'; Required = $false; Type = 'Text'}
            )
        }
    }
    
    # Validate all mapping groups
    [bool] ValidateAllMappings([object]$dataObject) {
        try {
            # Validate basic info
            $this.ValidateMappingGroup('BasicInfo', $dataObject)
            
            # Validate firewall
            $this.ValidateMappingGroup('Firewall', $dataObject)
            
            # Validate circuits
            $this.ValidateMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
            if ($dataObject.HasBackupCircuit) {
                $this.ValidateMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
            }
            
            # Validate VLANs
            $this.ValidateMappingGroup('VLANs', $dataObject.VLANs)
            
            return $true
        }
        catch {
            throw $_
        }
    }
    
    # Validate specific mapping group
    [void] ValidateMappingGroup([string]$groupName, [object]$dataObject) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            # Check required fields
            if ($mapping.Required) {
                $value = $dataObject.($mapping.Property)
                if ([string]::IsNullOrWhiteSpace($value)) {
                    throw "Required field missing: $($mapping.Property)"
                }
            }
            
            # Validate field format
            if ($mapping.ContainsKey('Validator')) {
                $value = $dataObject.($mapping.Property)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    if (-not $this.ValidateField($value, $mapping.Validator)) {
                        throw "Invalid $($mapping.Validator) format: $($mapping.Property) = $value"
                    }
                }
            }
        }
    }
    
    # Field validation
    [bool] ValidateField([string]$value, [string]$validatorType) {
        if ($validatorType -eq 'IP') {
            return [ValidationUtility]::ValidateIP($value)
        }
        return $true
    }
    
    # Set all mappings to UI
    [void] SetAllMappings([object]$dataObject) {
        $this.SetMappingGroup('BasicInfo', $dataObject)
        $this.SetMappingGroup('Firewall', $dataObject)
        $this.SetMappingGroup('VLANs', $dataObject.VLANs)
        $this.SetMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
        $this.SetMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
    }
    
    # Get all mappings from UI
    [void] GetAllMappings([object]$dataObject) {
        $this.GetMappingGroup('BasicInfo', $dataObject)
        $this.GetMappingGroup('Firewall', $dataObject)
        $this.GetMappingGroup('VLANs', $dataObject.VLANs)
        $this.GetMappingGroup('PrimaryCircuit', $dataObject.PrimaryCircuit)
        $this.GetMappingGroup('BackupCircuit', $dataObject.BackupCircuit)
    }
    
    # Clear all mappings
    [void] ClearAllMappings() {
        $this.ClearMappingGroup('BasicInfo')
        $this.ClearMappingGroup('Firewall')
        $this.ClearMappingGroup('VLANs')
        $this.ClearMappingGroup('PrimaryCircuit')
        $this.ClearMappingGroup('BackupCircuit')
    }
    
    # Set specific mapping group
    [void] SetMappingGroup([string]$groupName, [object]$dataObject) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            $control = $this.MainWindow.FindName($mapping.Control)
            if ($control) {
                $value = Get-SafeValue $dataObject.($mapping.Property)
                
                switch ($mapping.Type) {
                    'Text' { $control.Text = $value }
                    'CheckBox' { $control.IsChecked = [bool]$value }
                    'ComboBox' { $this.SetComboBoxSelection($control, $value) }
                }
            }
        }
    }
    
    # Get specific mapping group
[void] GetMappingGroup([string]$groupName, [object]$dataObject) {
    $group = $this.MappingGroups[$groupName]    
    foreach ($mapping in $group) {
        $control = $this.MainWindow.FindName($mapping.Control)
        if ($control) {
            $controlValue = ""
            switch ($mapping.Type) {
                'Text' { 
                    $controlValue = $control.Text.Trim()
                    $dataObject.($mapping.Property) = $controlValue
                }
                'CheckBox' { 
                    $controlValue = $control.IsChecked
                    $dataObject.($mapping.Property) = $controlValue
                }
                'ComboBox' { 
                    if ($control.SelectedItem) {
                        $controlValue = $control.SelectedItem.Content
                        $dataObject.($mapping.Property) = $controlValue
                    }
                }
            }
        } else {
        }
    }
}
    
    # Clear specific mapping group
    [void] ClearMappingGroup([string]$groupName) {
        $group = $this.MappingGroups[$groupName]
        foreach ($mapping in $group) {
            $control = $this.MainWindow.FindName($mapping.Control)
            if ($control) {
                switch ($mapping.Type) {
                    'Text' { $control.Text = "" }
                    'CheckBox' { $control.IsChecked = $false }
                    'ComboBox' { $control.SelectedIndex = -1 }
                }
            }
        }
    }
        
    # Helper method for ComboBox selection
    [void] SetComboBoxSelection([System.Windows.Controls.ComboBox]$ComboBox, [string]$Value) {
        Set-ComboBoxValue $ComboBox $Value -ByContent
    }
}

# ===================================================================
# DATA STORAGE CLASS
# ===================================================================

# Site data storage and persistence manager
class SiteDataStore {
    hidden [string]$DataFile = "$(Split-Path $PSScriptRoot -Parent)\Data\site_data.json"
    hidden [System.Collections.Generic.List[SiteEntry]]$Entries

    SiteDataStore() {
        $this.LoadData()
    }

    # Load site data from JSON file
    [void] LoadData() {   
    if (Test-Path $this.DataFile) {
            try {
                $jsonData = Get-Content $this.DataFile | ConvertFrom-Json
                $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
                foreach ($item in $jsonData) {
                    $site = [SiteEntry]::new()
                    $site.ID = $item.ID
                    
                    # Basic Info
                    $site.SiteCode = $item.SiteCode
                    $site.SiteSubnet = Get-SafeValue $item.SiteSubnet
                    $site.SiteSubnetCode = $item.SiteSubnetCode
                    $site.SiteName = $item.SiteName
                    $site.SiteAddress = $item.SiteAddress
                    $site.MainContactName = $item.MainContactName
                    $site.MainContactPhone = $item.MainContactPhone
                    $site.SecondContactName = $item.SecondContactName
                    $site.SecondContactPhone = $item.SecondContactPhone
                    
                    # Network Equipment
                    $site.SwitchCount = $item.SwitchCount
                    $site.Switches = [System.Collections.Generic.List[SwitchInfo]]::new()
                    if ($item.Switches) {
                        foreach ($switchItem in $item.Switches) {
                            $switch = [SwitchInfo]::new()
                            $switch.ManagementIP = Get-SafeValue $switchItem.ManagementIP
                            $switch.Name = Get-SafeValue $switchItem.Name
                            $switch.AssetTag = Get-SafeValue $switchItem.AssetTag
                            $switch.Version = Get-SafeValue $switchItem.Version
                            $switch.SerialNumber = Get-SafeValue $switchItem.SerialNumber
                            $site.Switches.Add($switch)
                        }
                    }

                    # Access Points
                    $site.APCount = if ($item.APCount) { $item.APCount } else { 0 }
                    $site.AccessPoints = [System.Collections.Generic.List[AccessPointInfo]]::new()
                    if ($item.AccessPoints) {
                        foreach ($apItem in $item.AccessPoints) {
                            $ap = [AccessPointInfo]::new()
                            $ap.ManagementIP = Get-SafeValue $apItem.ManagementIP
                            $ap.Name = Get-SafeValue $apItem.Name
                            $ap.AssetTag = Get-SafeValue $apItem.AssetTag
                            $ap.Version = Get-SafeValue $apItem.Version
                            $ap.SerialNumber = Get-SafeValue $apItem.SerialNumber
                            $site.AccessPoints.Add($ap)
                        }
                    }

                    # UPS
                    $site.UPSCount = if ($item.UPSCount) { $item.UPSCount } else { 0 }
                    $site.UPSDevices = [System.Collections.Generic.List[UPSInfo]]::new()
                    if ($item.UPSDevices) {
                        foreach ($upsItem in $item.UPSDevices) {
                            $ups = [UPSInfo]::new()
                            $ups.ManagementIP = Get-SafeValue $upsItem.ManagementIP
                            $ups.Name = Get-SafeValue $upsItem.Name
                            $ups.AssetTag = Get-SafeValue $upsItem.AssetTag
                            $ups.Version = Get-SafeValue $upsItem.Version
                            $ups.SerialNumber = Get-SafeValue $upsItem.SerialNumber
                            $site.UPSDevices.Add($ups)
                        }
                    }

                    # CCTV
                    $site.CCTVCount = if ($item.CCTVCount) { $item.CCTVCount } else { 0 }
                    $site.CCTVDevices = [System.Collections.Generic.List[CCTVInfo]]::new()
                    if ($item.CCTVDevices) {
                        foreach ($cctvItem in $item.CCTVDevices) {
                            $cctv = [CCTVInfo]::new()
                            $cctv.ManagementIP = Get-SafeValue $cctvItem.ManagementIP
                            $cctv.Name = Get-SafeValue $cctvItem.Name
                            $cctv.SerialNumber = Get-SafeValue $cctvItem.SerialNumber
                            $site.CCTVDevices.Add($cctv)
                        }
                    }

                    # Printer
                    $site.PrinterCount = if ($item.PrinterCount) { $item.PrinterCount } else { 0 }
                    $site.PrinterDevices = [System.Collections.Generic.List[PrinterInfo]]::new()
                    if ($item.PrinterDevices) {
                        foreach ($printerItem in $item.PrinterDevices) {
                            $printer = [PrinterInfo]::new()
                            $printer.ManagementIP = Get-SafeValue $printerItem.ManagementIP
                            $printer.Name = Get-SafeValue $printerItem.Name
                            $printer.Model = Get-SafeValue $printerItem.Model
                            $printer.SerialNumber = Get-SafeValue $printerItem.SerialNumber
                            $site.PrinterDevices.Add($printer)
                        }
                    }
                    
                    $site.FirewallIP = Get-SafeValue $item.FirewallIP
                    $site.FirewallName = Get-SafeValue $item.FirewallName
                    $site.FirewallVersion = Get-SafeValue $item.FirewallVersion
                    $site.FirewallSN = Get-SafeValue $item.FirewallSN
                    
                    # Circuits
                    if ($item.PrimaryCircuit) {
                        $site.PrimaryCircuit.Vendor = Get-SafeValue $item.PrimaryCircuit.Vendor
                        $site.PrimaryCircuit.CircuitType = Get-SafeValue $item.PrimaryCircuit.CircuitType
                        $site.PrimaryCircuit.CircuitID = Get-SafeValue $item.PrimaryCircuit.CircuitID
                        $site.PrimaryCircuit.DownloadSpeed = Get-SafeValue $item.PrimaryCircuit.DownloadSpeed
                        $site.PrimaryCircuit.UploadSpeed = Get-SafeValue $item.PrimaryCircuit.UploadSpeed
                        $site.PrimaryCircuit.IPAddress = Get-SafeValue $item.PrimaryCircuit.IPAddress
                        $site.PrimaryCircuit.SubnetMask = Get-SafeValue $item.PrimaryCircuit.SubnetMask
                        $site.PrimaryCircuit.DefaultGateway = Get-SafeValue $item.PrimaryCircuit.DefaultGateway
                        $site.PrimaryCircuit.DNS1 = Get-SafeValue $item.PrimaryCircuit.DNS1
                        $site.PrimaryCircuit.DNS2 = Get-SafeValue $item.PrimaryCircuit.DNS2
                        $site.PrimaryCircuit.RouterModel = Get-SafeValue $item.PrimaryCircuit.RouterModel
                        $site.PrimaryCircuit.RouterName = Get-SafeValue $item.PrimaryCircuit.RouterName
                        $site.PrimaryCircuit.RouterSN = Get-SafeValue $item.PrimaryCircuit.RouterSN
                        $site.PrimaryCircuit.PPPoEUsername = Get-SafeValue $item.PrimaryCircuit.PPPoEUsername
                        $site.PrimaryCircuit.PPPoEPassword = Get-SafeValue $item.PrimaryCircuit.PPPoEPassword
                        $site.PrimaryCircuit.HasModem = if ($item.PrimaryCircuit.HasModem) { $item.PrimaryCircuit.HasModem } else { $false }
                        $site.PrimaryCircuit.ModemModel = Get-SafeValue $item.PrimaryCircuit.ModemModel
                        $site.PrimaryCircuit.ModemName = Get-SafeValue $item.PrimaryCircuit.ModemName
                        $site.PrimaryCircuit.ModemSN = Get-SafeValue $item.PrimaryCircuit.ModemSN
                    }
                    
                    $site.HasBackupCircuit = if ($item.HasBackupCircuit) { $item.HasBackupCircuit } else { $false }
                    if ($item.BackupCircuit -and $site.HasBackupCircuit) {
                        $site.BackupCircuit.Vendor = Get-SafeValue $item.BackupCircuit.Vendor
                        $site.BackupCircuit.CircuitType = Get-SafeValue $item.BackupCircuit.CircuitType
                        $site.BackupCircuit.CircuitID = Get-SafeValue $item.BackupCircuit.CircuitID
                        $site.BackupCircuit.DownloadSpeed = Get-SafeValue $item.BackupCircuit.DownloadSpeed
                        $site.BackupCircuit.UploadSpeed = Get-SafeValue $item.BackupCircuit.UploadSpeed
                        $site.BackupCircuit.IPAddress = Get-SafeValue $item.BackupCircuit.IPAddress
                        $site.BackupCircuit.SubnetMask = Get-SafeValue $item.BackupCircuit.SubnetMask
                        $site.BackupCircuit.DefaultGateway = Get-SafeValue $item.BackupCircuit.DefaultGateway
                        $site.BackupCircuit.DNS1 = Get-SafeValue $item.BackupCircuit.DNS1
                        $site.BackupCircuit.DNS2 = Get-SafeValue $item.BackupCircuit.DNS2
                        $site.BackupCircuit.RouterModel = Get-SafeValue $item.BackupCircuit.RouterModel
                        $site.BackupCircuit.RouterName = Get-SafeValue $item.BackupCircuit.RouterName
                        $site.BackupCircuit.RouterSN = Get-SafeValue $item.BackupCircuit.RouterSN
                        $site.BackupCircuit.PPPoEUsername = Get-SafeValue $item.BackupCircuit.PPPoEUsername
                        $site.BackupCircuit.PPPoEPassword = Get-SafeValue $item.BackupCircuit.PPPoEPassword
                        $site.BackupCircuit.HasModem = if ($item.BackupCircuit.HasModem) { $item.BackupCircuit.HasModem } else { $false }
                        $site.BackupCircuit.ModemModel = Get-SafeValue $item.BackupCircuit.ModemModel
                        $site.BackupCircuit.ModemName = Get-SafeValue $item.BackupCircuit.ModemName
                        $site.BackupCircuit.ModemSN = Get-SafeValue $item.BackupCircuit.ModemSN
                    }
                    
                    # VLANs
                    if ($item.VLANs) {
                        $site.VLANs.VLAN100_Servers = Get-SafeValue $item.VLANs.VLAN100_Servers
                        $site.VLANs.VLAN101_NetworkDevices = Get-SafeValue $item.VLANs.VLAN101_NetworkDevices
                        $site.VLANs.VLAN102_UserDevices = Get-SafeValue $item.VLANs.VLAN102_UserDevices
                        $site.VLANs.VLAN103_UserDevices2 = Get-SafeValue $item.VLANs.VLAN103_UserDevices2
                        $site.VLANs.VLAN104_VOIP = Get-SafeValue $item.VLANs.VLAN104_VOIP
                        $site.VLANs.VLAN105_WiFiCorp = Get-SafeValue $item.VLANs.VLAN105_WiFiCorp
                        $site.VLANs.VLAN106_WiFiBYOD = Get-SafeValue $item.VLANs.VLAN106_WiFiBYOD
                        $site.VLANs.VLAN107_WiFiGuest = Get-SafeValue $item.VLANs.VLAN107_WiFiGuest
                        $site.VLANs.VLAN108_Spare = Get-SafeValue $item.VLANs.VLAN108_Spare
                        $site.VLANs.VLAN109_DMZ = Get-SafeValue $item.VLANs.VLAN109_DMZ
                        $site.VLANs.VLAN110_CCTV = Get-SafeValue $item.VLANs.VLAN110_CCTV
                    }
                    
                    $site.UpdateDisplayProperties()
                    $this.Entries.Add($site)
                }
            } catch {
                [System.Windows.MessageBox]::Show("Error loading site data: $_", "Data Load Error", "OK", "Warning")
                $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
                $this.SaveData()
            }
        } else {
            $this.Entries = [System.Collections.Generic.List[SiteEntry]]::new()
            $this.SaveData()
        }
    }

    # Save site data to JSON file
    [void] SaveData() {
    try {        
        if ($this.Entries.Count -eq 0) {
            if (Test-Path $this.DataFile) {
                Remove-Item $this.DataFile -Force
            }
        } else {
            $this.Entries | ConvertTo-Json -Depth 10 | Set-Content $this.DataFile
        }
    } catch {
        [System.Windows.MessageBox]::Show("Error saving site data: $_", "Data Save Error", "OK", "Error")
    }
}


    # Get all site entries
    [SiteEntry[]] GetAllEntries() {
        return $this.Entries.ToArray()
    }

    # Add new site entry with duplicate validation
    [bool] AddEntry([SiteEntry]$entry) {
    # Check for duplicate Site Code
    if ($this.Entries.SiteCode -contains $entry.SiteCode) {
        throw "Site code '$($entry.SiteCode)' already exists"
    }
    
    # Check for duplicate Site Subnet
    if ($this.Entries.SiteSubnet -contains $entry.SiteSubnet) {
        throw "Site subnet '$($entry.SiteSubnet)' already exists"
    }
    
    $entry.ID = $this.GetNextAvailableId()
    $entry.UpdateDisplayProperties()
    $this.Entries.Add($entry)
    $this.SaveData()
    return $true
    }

    # Update existing site entry
    [bool] UpdateEntry([SiteEntry]$entry) {
        $existingIndex = -1
        for ($i = 0; $i -lt $this.Entries.Count; $i++) {
            if ($this.Entries[$i].ID -eq $entry.ID) {
                $existingIndex = $i
                break
            }
        }
        
        if ($existingIndex -ge 0) {
            $entry.UpdateDisplayProperties()
            $this.Entries[$existingIndex] = $entry
            $this.SaveData()
            return $true
        }
        return $false
    }

    # Delete multiple site entries by ID
    [bool] DeleteEntries([int[]]$ids) {
        $countBefore = $this.Entries.Count
        $newEntries = [System.Collections.Generic.List[SiteEntry]]::new()
        
        foreach ($entry in $this.Entries) {
            if ($entry.ID -notin $ids) {
                $newEntries.Add($entry)
            }
        }
        
        $this.Entries = $newEntries
        
        if ($this.Entries.Count -lt $countBefore) {
            $this.SaveData()
            return $true
        }
        return $false
    }

    # Get next available ID for new entries
    hidden [int] GetNextAvailableId() {
        if ($this.Entries.Count -eq 0) { return 1 }
        $maxId = ($this.Entries.ID | Measure-Object -Maximum).Maximum
        for ($i = 1; $i -le $maxId; $i++) {
            if ($i -notin $this.Entries.ID) { return $i }
        }
        return $maxId + 1
    }
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
        [System.Windows.MessageBox]::Show("Error displaying site details: $_", "Display Error", "OK", "Error")
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
    [System.Windows.MessageBox]::Show("XAML file not found: $xamlFile", "File Error", "OK", "Error")
    exit
}

# Load and parse XAML
try {
    $xaml = Get-Content $xamlFile -Raw
    $xml = [xml]$xaml
}
catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", "OK", "Error")
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
    [System.Windows.MessageBox]::Show("Error creating window: $_", "Window Creation Error", "OK", "Error")
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

# ===================================================================
# EDIT SITE POPUP WINDOW FUNCTIONS
# ===================================================================
# Function to show the edit site window
function Show-EditSiteWindow {
    param([SiteEntry]$SiteToEdit)
    
    if (-not $SiteToEdit) {
        Show-CustomDialog "No site selected for editing." "Selection Required" "OK" "Warning"
        return $false
    }
    
    try {
        # Load the Edit Site XAML
        $editXamlFile = Join-Path (Split-Path $scriptPath -Parent) "UI" | Join-Path -ChildPath "EditSiteWindow.xaml"
        
        if (-not (Test-Path $editXamlFile)) {
            Show-CustomDialog "Edit window XAML file not found: $editXamlFile" "File Error" "OK" "Error"
            return $false
        }
        
        $editXaml = Get-Content $editXamlFile -Raw
        $editXml = [xml]$editXaml
        $editReader = New-Object System.Xml.XmlNodeReader $editXml
        $editWindow = [Windows.Markup.XamlReader]::Load($editReader)
        
        # Set window properties
        $editWindow.Owner = $mainWin
        $editWindow.Title = "Edit Site: $($SiteToEdit.SiteCode)"
        
        # Initialize edit window managers
        $editDeviceManager = [EditDevicePanelManager]::new($editWindow)
        $editFieldManager = [FieldMappingManager]::new($editWindow)
        
        # Get all edit window controls
        $editControls = Get-EditWindowControls -EditWindow $editWindow
        
        # Set up event handlers for the edit window
        Setup-EditWindowEventHandlers -EditWindow $editWindow -EditControls $editControls -EditDeviceManager $editDeviceManager
        
        # Populate the edit window with existing site data
        Populate-EditWindow -SiteToEdit $SiteToEdit -EditControls $editControls -EditDeviceManager $editDeviceManager -EditFieldManager $editFieldManager
        
        # Setup button event handlers
        $editControls.btnSaveChanges.Add_Click({
            if (Save-EditedSite -SiteToEdit $SiteToEdit -EditControls $editControls -EditDeviceManager $editDeviceManager -EditFieldManager $editFieldManager) {
                $editWindow.DialogResult = $true
                $editWindow.Close()
            }
        })
        
        $editControls.btnCancelEdit.Add_Click({
            $editWindow.DialogResult = $false
            $editWindow.Close()
        })
        
        $editControls.btnResetForm.Add_Click({
            Populate-EditWindow -SiteToEdit $SiteToEdit -EditControls $editControls -EditDeviceManager $editDeviceManager -EditFieldManager $editFieldManager
            $editControls.txtEditStatus.Text = "Form reset to original values"
            $editControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Blue
        })
        
        # Show the window and return result
        $result = $editWindow.ShowDialog()
        
        if ($result -eq $true) {
            # Refresh the main data grid
            Update-DataGridWithSearch
            Show-ValidationError "Site '$($SiteToEdit.SiteCode)' updated successfully!" "Success"
            return $true
        }
        
        return $false
        
    } catch {
        Show-CustomDialog "Error opening edit window: $_" "Edit Window Error" "OK" "Error"
        return $false
    }
}

# Function to get all edit window control references
function Get-EditWindowControls {
    param([object]$EditWindow)
    
    $controls = @{}
    
    # Basic Info controls
    $controls.txtEditSiteCode = $EditWindow.FindName("txtEditSiteCode")
    $controls.txtEditSiteSubnet = $EditWindow.FindName("txtEditSiteSubnet")
    $controls.txtEditSiteSubnetCode = $EditWindow.FindName("txtEditSiteSubnetCode")
    $controls.txtEditSiteName = $EditWindow.FindName("txtEditSiteName")
    $controls.txtEditSiteAddress = $EditWindow.FindName("txtEditSiteAddress")
    $controls.txtEditMainContactName = $EditWindow.FindName("txtEditMainContactName")
    $controls.txtEditMainContactPhone = $EditWindow.FindName("txtEditMainContactPhone")
    $controls.txtEditSecondContactName = $EditWindow.FindName("txtEditSecondContactName")
    $controls.txtEditSecondContactPhone = $EditWindow.FindName("txtEditSecondContactPhone")
    
    # Device controls
    $controls.cmbEditSwitchCount = $EditWindow.FindName("cmbEditSwitchCount")
    $controls.stkEditSwitches = $EditWindow.FindName("stkEditSwitches")
    $controls.cmbEditAPCount = $EditWindow.FindName("cmbEditAPCount")
    $controls.stkEditAccessPoints = $EditWindow.FindName("stkEditAccessPoints")
    $controls.cmbEditUPSCount = $EditWindow.FindName("cmbEditUPSCount")
    $controls.stkEditUPS = $EditWindow.FindName("stkEditUPS")
    $controls.cmbEditCCTVCount = $EditWindow.FindName("cmbEditCCTVCount")
    $controls.stkEditCCTV = $EditWindow.FindName("stkEditCCTV")
    $controls.cmbEditPrinterCount = $EditWindow.FindName("cmbEditPrinterCount")
    $controls.stkEditPrinter = $EditWindow.FindName("stkEditPrinter")
    
    # Firewall controls
    $controls.txtEditFirewallIP = $EditWindow.FindName("txtEditFirewallIP")
    $controls.txtEditFirewallName = $EditWindow.FindName("txtEditFirewallName")
    $controls.txtEditFirewallVersion = $EditWindow.FindName("txtEditFirewallVersion")
    $controls.txtEditFirewallSN = $EditWindow.FindName("txtEditFirewallSN")
    
    # Primary Circuit controls
    $controls.txtEditPrimaryVendor = $EditWindow.FindName("txtEditPrimaryVendor")
    $controls.cmbEditPrimaryCircuitType = $EditWindow.FindName("cmbEditPrimaryCircuitType")
    $controls.stkEditPrimaryGPON = $EditWindow.FindName("stkEditPrimaryGPON")
    $controls.txtEditPrimaryPPPoEUsername = $EditWindow.FindName("txtEditPrimaryPPPoEUsername")
    $controls.txtEditPrimaryPPPoEPassword = $EditWindow.FindName("txtEditPrimaryPPPoEPassword")
    $controls.txtEditPrimaryCircuitID = $EditWindow.FindName("txtEditPrimaryCircuitID")
    $controls.txtEditPrimaryDownloadSpeed = $EditWindow.FindName("txtEditPrimaryDownloadSpeed")
    $controls.txtEditPrimaryUploadSpeed = $EditWindow.FindName("txtEditPrimaryUploadSpeed")
    $controls.txtEditPrimaryIPAddress = $EditWindow.FindName("txtEditPrimaryIPAddress")
    $controls.txtEditPrimarySubnetMask = $EditWindow.FindName("txtEditPrimarySubnetMask")
    $controls.txtEditPrimaryDefaultGateway = $EditWindow.FindName("txtEditPrimaryDefaultGateway")
    $controls.txtEditPrimaryDNS1 = $EditWindow.FindName("txtEditPrimaryDNS1")
    $controls.txtEditPrimaryDNS2 = $EditWindow.FindName("txtEditPrimaryDNS2")
    $controls.txtEditPrimaryRouterModel = $EditWindow.FindName("txtEditPrimaryRouterModel")
    $controls.txtEditPrimaryRouterName = $EditWindow.FindName("txtEditPrimaryRouterName")
    $controls.txtEditPrimaryRouterSN = $EditWindow.FindName("txtEditPrimaryRouterSN")
    $controls.chkEditPrimaryHasModem = $EditWindow.FindName("chkEditPrimaryHasModem")
    $controls.stkEditPrimaryModem = $EditWindow.FindName("stkEditPrimaryModem")
    $controls.txtEditPrimaryModemModel = $EditWindow.FindName("txtEditPrimaryModemModel")
    $controls.txtEditPrimaryModemName = $EditWindow.FindName("txtEditPrimaryModemName")
    $controls.txtEditPrimaryModemSN = $EditWindow.FindName("txtEditPrimaryModemSN")
    
    # Backup Circuit controls
    $controls.chkEditHasBackupCircuit = $EditWindow.FindName("chkEditHasBackupCircuit")
    $controls.grdEditBackupCircuit = $EditWindow.FindName("grdEditBackupCircuit")
    $controls.txtEditBackupVendor = $EditWindow.FindName("txtEditBackupVendor")
    $controls.cmbEditBackupCircuitType = $EditWindow.FindName("cmbEditBackupCircuitType")
    $controls.stkEditBackupGPON = $EditWindow.FindName("stkEditBackupGPON")
    $controls.txtEditBackupPPPoEUsername = $EditWindow.FindName("txtEditBackupPPPoEUsername")
    $controls.txtEditBackupPPPoEPassword = $EditWindow.FindName("txtEditBackupPPPoEPassword")
    $controls.txtEditBackupCircuitID = $EditWindow.FindName("txtEditBackupCircuitID")
    $controls.txtEditBackupDownloadSpeed = $EditWindow.FindName("txtEditBackupDownloadSpeed")
    $controls.txtEditBackupUploadSpeed = $EditWindow.FindName("txtEditBackupUploadSpeed")
    $controls.txtEditBackupIPAddress = $EditWindow.FindName("txtEditBackupIPAddress")
    $controls.txtEditBackupSubnetMask = $EditWindow.FindName("txtEditBackupSubnetMask")
    $controls.txtEditBackupDefaultGateway = $EditWindow.FindName("txtEditBackupDefaultGateway")
    $controls.txtEditBackupDNS1 = $EditWindow.FindName("txtEditBackupDNS1")
    $controls.txtEditBackupDNS2 = $EditWindow.FindName("txtEditBackupDNS2")
    $controls.txtEditBackupRouterModel = $EditWindow.FindName("txtEditBackupRouterModel")
    $controls.txtEditBackupRouterName = $EditWindow.FindName("txtEditBackupRouterName")
    $controls.txtEditBackupRouterSN = $EditWindow.FindName("txtEditBackupRouterSN")
    $controls.chkEditBackupHasModem = $EditWindow.FindName("chkEditBackupHasModem")
    $controls.stkEditBackupModem = $EditWindow.FindName("stkEditBackupModem")
    $controls.txtEditBackupModemModel = $EditWindow.FindName("txtEditBackupModemModel")
    $controls.txtEditBackupModemName = $EditWindow.FindName("txtEditBackupModemName")
    $controls.txtEditBackupModemSN = $EditWindow.FindName("txtEditBackupModemSN")
    
    # VLAN controls
    $controls.txtEditVlan100 = $EditWindow.FindName("txtEditVlan100")
    $controls.txtEditVlan101 = $EditWindow.FindName("txtEditVlan101")
    $controls.txtEditVlan102 = $EditWindow.FindName("txtEditVlan102")
    $controls.txtEditVlan103 = $EditWindow.FindName("txtEditVlan103")
    $controls.txtEditVlan104 = $EditWindow.FindName("txtEditVlan104")
    $controls.txtEditVlan105 = $EditWindow.FindName("txtEditVlan105")
    $controls.txtEditVlan106 = $EditWindow.FindName("txtEditVlan106")
    $controls.txtEditVlan107 = $EditWindow.FindName("txtEditVlan107")
    $controls.txtEditVlan108 = $EditWindow.FindName("txtEditVlan108")
    $controls.txtEditVlan109 = $EditWindow.FindName("txtEditVlan109")
    $controls.txtEditVlan110 = $EditWindow.FindName("txtEditVlan110")
    
    # Button and status controls
    $controls.btnSaveChanges = $EditWindow.FindName("btnSaveChanges")
    $controls.btnCancelEdit = $EditWindow.FindName("btnCancelEdit")
    $controls.btnResetForm = $EditWindow.FindName("btnResetForm")
    $controls.txtEditStatus = $EditWindow.FindName("txtEditStatus")
    
    return $controls
}

# Function to set up event handlers for the edit window
function Setup-EditWindowEventHandlers {
    param(
        [object]$EditWindow,
        [hashtable]$EditControls,
        [object]$EditDeviceManager
    )
    
    # Device count change handlers
    $EditControls.cmbEditSwitchCount.Add_SelectionChanged({
        if ($EditControls.cmbEditSwitchCount.SelectedItem) {
            $count = [int]$EditControls.cmbEditSwitchCount.SelectedItem.Content
            $EditDeviceManager.UpdateDevicePanels('Switch', $count)
            
            # Auto-populate after panels are created
            if ($count -gt 0) {
                $siteCode = $EditControls.txtEditSiteCode.Text
                $siteSubnet = $EditControls.txtEditSiteSubnet.Text
                
                if (-not [string]::IsNullOrWhiteSpace($siteCode)) {
                    $EditDeviceManager.UpdateDeviceNamesFromSiteCode('Switch', $siteCode)
                }
                if (-not [string]::IsNullOrWhiteSpace($siteSubnet) -and $siteSubnet -match '^(\d+\.\d+)\.') {
                    $EditDeviceManager.UpdateDeviceIPsFromSubnet('Switch', $matches[1])
                }
            }
        }
    })
    
    $EditControls.cmbEditAPCount.Add_SelectionChanged({
        if ($EditControls.cmbEditAPCount.SelectedItem) {
            $count = [int]$EditControls.cmbEditAPCount.SelectedItem.Content
            $EditDeviceManager.UpdateDevicePanels('AccessPoint', $count)
            
            # Auto-populate after panels are created
            if ($count -gt 0) {
                $siteCode = $EditControls.txtEditSiteCode.Text
                $siteSubnet = $EditControls.txtEditSiteSubnet.Text
                
                if (-not [string]::IsNullOrWhiteSpace($siteCode)) {
                    $EditDeviceManager.UpdateDeviceNamesFromSiteCode('AccessPoint', $siteCode)
                }
                if (-not [string]::IsNullOrWhiteSpace($siteSubnet) -and $siteSubnet -match '^(\d+\.\d+)\.') {
                    $EditDeviceManager.UpdateDeviceIPsFromSubnet('AccessPoint', $matches[1])
                }
            }
        }
    })
    
    $EditControls.cmbEditUPSCount.Add_SelectionChanged({
        if ($EditControls.cmbEditUPSCount.SelectedItem) {
            $count = [int]$EditControls.cmbEditUPSCount.SelectedItem.Content
            $EditDeviceManager.UpdateDevicePanels('UPS', $count)
            
            # Auto-populate after panels are created
            if ($count -gt 0) {
                $siteCode = $EditControls.txtEditSiteCode.Text
                $siteSubnet = $EditControls.txtEditSiteSubnet.Text
                
                if (-not [string]::IsNullOrWhiteSpace($siteCode)) {
                    $EditDeviceManager.UpdateDeviceNamesFromSiteCode('UPS', $siteCode)
                }
                if (-not [string]::IsNullOrWhiteSpace($siteSubnet) -and $siteSubnet -match '^(\d+\.\d+)\.') {
                    $EditDeviceManager.UpdateDeviceIPsFromSubnet('UPS', $matches[1])
                }
            }
        }
    })
    
    $EditControls.cmbEditCCTVCount.Add_SelectionChanged({
        if ($EditControls.cmbEditCCTVCount.SelectedItem) {
            $count = [int]$EditControls.cmbEditCCTVCount.SelectedItem.Content
            $EditDeviceManager.UpdateDevicePanels('CCTV', $count)
            
            # Auto-populate after panels are created
            if ($count -gt 0) {
                $siteCode = $EditControls.txtEditSiteCode.Text
                $siteSubnet = $EditControls.txtEditSiteSubnet.Text
                
                if (-not [string]::IsNullOrWhiteSpace($siteCode)) {
                    $EditDeviceManager.UpdateDeviceNamesFromSiteCode('CCTV', $siteCode)
                }
                if (-not [string]::IsNullOrWhiteSpace($siteSubnet) -and $siteSubnet -match '^(\d+\.\d+)\.') {
                    $EditDeviceManager.UpdateDeviceIPsFromSubnet('CCTV', $matches[1])
                }
            }
        }
    })

    $EditControls.cmbEditPrinterCount.Add_SelectionChanged({
        if ($EditControls.cmbEditPrinterCount.SelectedItem) {
            $count = [int]$EditControls.cmbEditPrinterCount.SelectedItem.Content
            $EditDeviceManager.UpdateDevicePanels('Printer', $count)
            
            # Auto-populate after panels are created
            if ($count -gt 0) {
                $siteCode = $EditControls.txtEditSiteCode.Text
                $siteSubnet = $EditControls.txtEditSiteSubnet.Text
                
                if (-not [string]::IsNullOrWhiteSpace($siteCode)) {
                    $EditDeviceManager.UpdateDeviceNamesFromSiteCode('Printer', $siteCode)
                }
                if (-not [string]::IsNullOrWhiteSpace($siteSubnet) -and $siteSubnet -match '^(\d+\.\d+)\.') {
                    $EditDeviceManager.UpdateDeviceIPsFromSubnet('Printer', $matches[1])
                }
            }
        }
    })
    
    # Backup circuit checkbox
    $EditControls.chkEditHasBackupCircuit.Add_Checked({
        if ($EditControls.grdEditBackupCircuit) {
            $EditControls.grdEditBackupCircuit.Visibility = "Visible"
        }
    })
    
    $EditControls.chkEditHasBackupCircuit.Add_Unchecked({
        if ($EditControls.grdEditBackupCircuit) {
            $EditControls.grdEditBackupCircuit.Visibility = "Collapsed"
        }
    })
    
    # Primary modem checkbox
    $EditControls.chkEditPrimaryHasModem.Add_Checked({
        if ($EditControls.stkEditPrimaryModem) {
            $EditControls.stkEditPrimaryModem.Visibility = "Visible"
        }
    })
    
    $EditControls.chkEditPrimaryHasModem.Add_Unchecked({
        if ($EditControls.stkEditPrimaryModem) {
            $EditControls.stkEditPrimaryModem.Visibility = "Collapsed"
        }
    })
    
    # Backup modem checkbox
    $EditControls.chkEditBackupHasModem.Add_Checked({
        if ($EditControls.stkEditBackupModem) {
            $EditControls.stkEditBackupModem.Visibility = "Visible"
        }
    })
    
    $EditControls.chkEditBackupHasModem.Add_Unchecked({
        if ($EditControls.stkEditBackupModem) {
            $EditControls.stkEditBackupModem.Visibility = "Collapsed"
        }
    })
    
    # Primary circuit type selection changed
    $EditControls.cmbEditPrimaryCircuitType.Add_SelectionChanged({
        if ($EditControls.stkEditPrimaryGPON) {
            if ($EditControls.cmbEditPrimaryCircuitType.SelectedItem -and $EditControls.cmbEditPrimaryCircuitType.SelectedItem.Content -eq "GPON Fiber") {
                $EditControls.stkEditPrimaryGPON.Visibility = "Visible"
            } else {
                $EditControls.stkEditPrimaryGPON.Visibility = "Collapsed"
            }
        }
    })
    
    # Backup circuit type selection changed
    $EditControls.cmbEditBackupCircuitType.Add_SelectionChanged({
        if ($EditControls.stkEditBackupGPON) {
            if ($EditControls.cmbEditBackupCircuitType.SelectedItem -and $EditControls.cmbEditBackupCircuitType.SelectedItem.Content -eq "GPON Fiber") {
                $EditControls.stkEditBackupGPON.Visibility = "Visible"
            } else {
                $EditControls.stkEditBackupGPON.Visibility = "Collapsed"
            }
        }
    })

    # Site Code auto-population using centralized function
    $EditControls.txtEditSiteCode.Add_TextChanged({
        Update-DeviceNamesFromSiteCode -SiteCode $EditControls.txtEditSiteCode.Text -DeviceManager $EditDeviceManager -FirewallNameControl $EditControls.txtEditFirewallName
    })

    # Site Subnet auto-population using centralized function
    $EditControls.txtEditSiteSubnet.Add_TextChanged({
        $editVlanControls = @{
            VLAN100 = $EditControls.txtEditVlan100
            VLAN101 = $EditControls.txtEditVlan101
            VLAN102 = $EditControls.txtEditVlan102
            VLAN103 = $EditControls.txtEditVlan103
            VLAN104 = $EditControls.txtEditVlan104
            VLAN105 = $EditControls.txtEditVlan105
            VLAN106 = $EditControls.txtEditVlan106
            VLAN107 = $EditControls.txtEditVlan107
            VLAN108 = $EditControls.txtEditVlan108
            VLAN109 = $EditControls.txtEditVlan109
            VLAN110 = $EditControls.txtEditVlan110
        }
        Update-VLANsAndIPsFromSubnet -SubnetInput $EditControls.txtEditSiteSubnet.Text -VLANControls $editVlanControls -DeviceManager $EditDeviceManager -FirewallIPControl $EditControls.txtEditFirewallIP -SiteSubnetCodeControl $EditControls.txtEditSiteSubnetCode
    })

}

# Function to populate the edit window with existing site data
function Populate-EditWindow {
    param(
        [SiteEntry]$SiteToEdit,
        [hashtable]$EditControls,
        [object]$EditDeviceManager,
        [object]$EditFieldManager
    )
    
    try {
        
        # Basic Info
        $EditControls.txtEditSiteCode.Text = $SiteToEdit.SiteCode
        $EditControls.txtEditSiteSubnet.Text = $SiteToEdit.SiteSubnet
        $EditControls.txtEditSiteSubnetCode.Text = $SiteToEdit.SiteSubnetCode
        $EditControls.txtEditSiteName.Text = $SiteToEdit.SiteName
        $EditControls.txtEditSiteAddress.Text = $SiteToEdit.SiteAddress
        $EditControls.txtEditMainContactName.Text = $SiteToEdit.MainContactName
        $EditControls.txtEditMainContactPhone.Text = $SiteToEdit.MainContactPhone
        $EditControls.txtEditSecondContactName.Text = $SiteToEdit.SecondContactName
        $EditControls.txtEditSecondContactPhone.Text = $SiteToEdit.SecondContactPhone
        
        # Firewall
        $EditControls.txtEditFirewallIP.Text = $SiteToEdit.FirewallIP
        $EditControls.txtEditFirewallName.Text = $SiteToEdit.FirewallName
        $EditControls.txtEditFirewallVersion.Text = $SiteToEdit.FirewallVersion
        $EditControls.txtEditFirewallSN.Text = $SiteToEdit.FirewallSN
        
        # Primary Circuit
        $EditControls.txtEditPrimaryVendor.Text = $SiteToEdit.PrimaryCircuit.Vendor
        Set-ComboBoxValue $EditControls.cmbEditPrimaryCircuitType $SiteToEdit.PrimaryCircuit.CircuitType -ByContent
        $EditControls.txtEditPrimaryPPPoEUsername.Text = $SiteToEdit.PrimaryCircuit.PPPoEUsername
        $EditControls.txtEditPrimaryPPPoEPassword.Text = $SiteToEdit.PrimaryCircuit.PPPoEPassword
        $EditControls.txtEditPrimaryCircuitID.Text = $SiteToEdit.PrimaryCircuit.CircuitID
        $EditControls.txtEditPrimaryDownloadSpeed.Text = $SiteToEdit.PrimaryCircuit.DownloadSpeed
        $EditControls.txtEditPrimaryUploadSpeed.Text = $SiteToEdit.PrimaryCircuit.UploadSpeed
        $EditControls.txtEditPrimaryIPAddress.Text = $SiteToEdit.PrimaryCircuit.IPAddress
        $EditControls.txtEditPrimarySubnetMask.Text = $SiteToEdit.PrimaryCircuit.SubnetMask
        $EditControls.txtEditPrimaryDefaultGateway.Text = $SiteToEdit.PrimaryCircuit.DefaultGateway
        $EditControls.txtEditPrimaryDNS1.Text = $SiteToEdit.PrimaryCircuit.DNS1
        $EditControls.txtEditPrimaryDNS2.Text = $SiteToEdit.PrimaryCircuit.DNS2
        $EditControls.txtEditPrimaryRouterModel.Text = $SiteToEdit.PrimaryCircuit.RouterModel
        $EditControls.txtEditPrimaryRouterName.Text = $SiteToEdit.PrimaryCircuit.RouterName
        $EditControls.txtEditPrimaryRouterSN.Text = $SiteToEdit.PrimaryCircuit.RouterSN
        $EditControls.chkEditPrimaryHasModem.IsChecked = $SiteToEdit.PrimaryCircuit.HasModem
        $EditControls.txtEditPrimaryModemModel.Text = $SiteToEdit.PrimaryCircuit.ModemModel
        $EditControls.txtEditPrimaryModemName.Text = $SiteToEdit.PrimaryCircuit.ModemName
        $EditControls.txtEditPrimaryModemSN.Text = $SiteToEdit.PrimaryCircuit.ModemSN
        
        # Backup Circuit
        $EditControls.chkEditHasBackupCircuit.IsChecked = $SiteToEdit.HasBackupCircuit
        if ($SiteToEdit.HasBackupCircuit) {
        $EditControls.txtEditBackupVendor.Text = $SiteToEdit.BackupCircuit.Vendor
        Set-ComboBoxValue $EditControls.cmbEditBackupCircuitType $SiteToEdit.BackupCircuit.CircuitType -ByContent
        $EditControls.txtEditBackupPPPoEUsername.Text = $SiteToEdit.BackupCircuit.PPPoEUsername
        $EditControls.txtEditBackupPPPoEPassword.Text = $SiteToEdit.BackupCircuit.PPPoEPassword
        $EditControls.txtEditBackupCircuitID.Text = $SiteToEdit.BackupCircuit.CircuitID
        $EditControls.txtEditBackupDownloadSpeed.Text = $SiteToEdit.BackupCircuit.DownloadSpeed
        $EditControls.txtEditBackupUploadSpeed.Text = $SiteToEdit.BackupCircuit.UploadSpeed
        $EditControls.txtEditBackupIPAddress.Text = $SiteToEdit.BackupCircuit.IPAddress
        $EditControls.txtEditBackupSubnetMask.Text = $SiteToEdit.BackupCircuit.SubnetMask
        $EditControls.txtEditBackupDefaultGateway.Text = $SiteToEdit.BackupCircuit.DefaultGateway
        $EditControls.txtEditBackupDNS1.Text = $SiteToEdit.BackupCircuit.DNS1
        $EditControls.txtEditBackupDNS2.Text = $SiteToEdit.BackupCircuit.DNS2
        $EditControls.txtEditBackupRouterModel.Text = $SiteToEdit.BackupCircuit.RouterModel
        $EditControls.txtEditBackupRouterName.Text = $SiteToEdit.BackupCircuit.RouterName
        $EditControls.txtEditBackupRouterSN.Text = $SiteToEdit.BackupCircuit.RouterSN
        $EditControls.chkEditBackupHasModem.IsChecked = $SiteToEdit.BackupCircuit.HasModem
        $EditControls.txtEditBackupModemModel.Text = $SiteToEdit.BackupCircuit.ModemModel
        $EditControls.txtEditBackupModemName.Text = $SiteToEdit.BackupCircuit.ModemName
        $EditControls.txtEditBackupModemSN.Text = $SiteToEdit.BackupCircuit.ModemSN
        }
        
        # VLANs
        $EditControls.txtEditVlan100.Text = $SiteToEdit.VLANs.VLAN100_Servers
        $EditControls.txtEditVlan101.Text = $SiteToEdit.VLANs.VLAN101_NetworkDevices
        $EditControls.txtEditVlan102.Text = $SiteToEdit.VLANs.VLAN102_UserDevices
        $EditControls.txtEditVlan103.Text = $SiteToEdit.VLANs.VLAN103_UserDevices2
        $EditControls.txtEditVlan104.Text = $SiteToEdit.VLANs.VLAN104_VOIP
        $EditControls.txtEditVlan105.Text = $SiteToEdit.VLANs.VLAN105_WiFiCorp
        $EditControls.txtEditVlan106.Text = $SiteToEdit.VLANs.VLAN106_WiFiBYOD
        $EditControls.txtEditVlan107.Text = $SiteToEdit.VLANs.VLAN107_WiFiGuest
        $EditControls.txtEditVlan108.Text = $SiteToEdit.VLANs.VLAN108_Spare
        $EditControls.txtEditVlan109.Text = $SiteToEdit.VLANs.VLAN109_DMZ
        $EditControls.txtEditVlan110.Text = $SiteToEdit.VLANs.VLAN110_CCTV
        
        # Set device counts and populate device panels
        Set-ComboBoxValue $EditControls.cmbEditSwitchCount $SiteToEdit.SwitchCount -ByContent
        $EditDeviceManager.UpdateDevicePanels('Switch', $SiteToEdit.SwitchCount)
        [UniversalDataCollector]::PopulateDevicePanels($EditDeviceManager.Configurations['Switch'], $EditDeviceManager.StackPanels['Switch'], $SiteToEdit.Switches, "txtEdit")
        
        Set-ComboBoxValue $EditControls.cmbEditAPCount $SiteToEdit.APCount -ByContent
        $EditDeviceManager.UpdateDevicePanels('AccessPoint', $SiteToEdit.APCount)
        [UniversalDataCollector]::PopulateDevicePanels($EditDeviceManager.Configurations['AccessPoint'], $EditDeviceManager.StackPanels['AccessPoint'], $SiteToEdit.AccessPoints, "txtEdit")
        
        Set-ComboBoxValue $EditControls.cmbEditUPSCount $SiteToEdit.UPSCount -ByContent
        $EditDeviceManager.UpdateDevicePanels('UPS', $SiteToEdit.UPSCount)
        [UniversalDataCollector]::PopulateDevicePanels($EditDeviceManager.Configurations['UPS'], $EditDeviceManager.StackPanels['UPS'], $SiteToEdit.UPSDevices, "txtEdit")
        
        Set-ComboBoxValue $EditControls.cmbEditCCTVCount $SiteToEdit.CCTVCount -ByContent
        $EditDeviceManager.UpdateDevicePanels('CCTV', $SiteToEdit.CCTVCount)
        [UniversalDataCollector]::PopulateDevicePanels($EditDeviceManager.Configurations['CCTV'], $EditDeviceManager.StackPanels['CCTV'], $SiteToEdit.CCTVDevices, "txtEdit")

        Set-ComboBoxValue $EditControls.cmbEditPrinterCount $SiteToEdit.PrinterCount -ByContent
        $EditDeviceManager.UpdateDevicePanels('Printer', $SiteToEdit.PrinterCount)
        [UniversalDataCollector]::PopulateDevicePanels($EditDeviceManager.Configurations['Printer'], $EditDeviceManager.StackPanels['Printer'], $SiteToEdit.PrinterDevices, "txtEdit")
        
        # Set initial visibility states
        if ($SiteToEdit.HasBackupCircuit) {
            $EditControls.grdEditBackupCircuit.Visibility = "Visible"
        } else {
            $EditControls.grdEditBackupCircuit.Visibility = "Collapsed"
        }
        
        if ($SiteToEdit.PrimaryCircuit.HasModem) {
            $EditControls.stkEditPrimaryModem.Visibility = "Visible"
        } else {
            $EditControls.stkEditPrimaryModem.Visibility = "Collapsed"
        }
        
        if ($SiteToEdit.BackupCircuit.HasModem) {
            $EditControls.stkEditBackupModem.Visibility = "Visible"
        } else {
            $EditControls.stkEditBackupModem.Visibility = "Collapsed"
        }
        
        # Set GPON visibility based on circuit types
        if ($SiteToEdit.PrimaryCircuit.CircuitType -eq "GPON Fiber") {
            $EditControls.stkEditPrimaryGPON.Visibility = "Visible"
        } else {
            $EditControls.stkEditPrimaryGPON.Visibility = "Collapsed"
        }
        
        if ($SiteToEdit.BackupCircuit.CircuitType -eq "GPON Fiber") {
            $EditControls.stkEditBackupGPON.Visibility = "Visible"
        } else {
            $EditControls.stkEditBackupGPON.Visibility = "Collapsed"
        }
        
        $EditControls.txtEditStatus.Text = "Site data loaded successfully"
        $EditControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Green
        
    } catch {
        $EditControls.txtEditStatus.Text = "Error loading site data: $_"
        $EditControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
}

# Function to save the edited site data
function Save-EditedSite {
    param(
        [SiteEntry]$SiteToEdit,
        [hashtable]$EditControls,
        [object]$EditDeviceManager,
        [object]$EditFieldManager
    )
    
    try {
        # Create a copy of the site to edit
        $editedSite = [SiteEntry]::new()
        $editedSite.ID = $SiteToEdit.ID  # Keep the same ID
        
        # Basic Info
        $editedSite.SiteCode = $EditControls.txtEditSiteCode.Text.Trim()
        $editedSite.SiteSubnet = $EditControls.txtEditSiteSubnet.Text.Trim()
        $editedSite.SiteSubnetCode = $EditControls.txtEditSiteSubnetCode.Text.Trim()
        $editedSite.SiteName = $EditControls.txtEditSiteName.Text.Trim()
        $editedSite.SiteAddress = $EditControls.txtEditSiteAddress.Text.Trim()
        $editedSite.MainContactName = $EditControls.txtEditMainContactName.Text.Trim()
        $editedSite.MainContactPhone = $EditControls.txtEditMainContactPhone.Text.Trim()
        $editedSite.SecondContactName = $EditControls.txtEditSecondContactName.Text.Trim()
        $editedSite.SecondContactPhone = $EditControls.txtEditSecondContactPhone.Text.Trim()
        
        # Use centralized validation (exclude current site from duplicate checks)
        try {
            Validate-SiteBasicInfo -Site $editedSite -StatusControl $EditControls.txtEditStatus -ExcludeSiteID $editedSite.ID
        } catch {
            return $false
        }
        
        # Firewall
        $editedSite.FirewallIP = $EditControls.txtEditFirewallIP.Text.Trim()
        $editedSite.FirewallName = $EditControls.txtEditFirewallName.Text.Trim()
        $editedSite.FirewallVersion = $EditControls.txtEditFirewallVersion.Text.Trim()
        $editedSite.FirewallSN = $EditControls.txtEditFirewallSN.Text.Trim()
        
        # Primary Circuit
        $editedSite.PrimaryCircuit.Vendor = $EditControls.txtEditPrimaryVendor.Text.Trim()
        if ($EditControls.cmbEditPrimaryCircuitType.SelectedItem) {
            $editedSite.PrimaryCircuit.CircuitType = $EditControls.cmbEditPrimaryCircuitType.SelectedItem.Content
        }
        $editedSite.PrimaryCircuit.PPPoEUsername = $EditControls.txtEditPrimaryPPPoEUsername.Text.Trim()
        $editedSite.PrimaryCircuit.PPPoEPassword = $EditControls.txtEditPrimaryPPPoEPassword.Text.Trim()
        $editedSite.PrimaryCircuit.CircuitID = $EditControls.txtEditPrimaryCircuitID.Text.Trim()
        $editedSite.PrimaryCircuit.DownloadSpeed = $EditControls.txtEditPrimaryDownloadSpeed.Text.Trim()
        $editedSite.PrimaryCircuit.UploadSpeed = $EditControls.txtEditPrimaryUploadSpeed.Text.Trim()
        $editedSite.PrimaryCircuit.IPAddress = $EditControls.txtEditPrimaryIPAddress.Text.Trim()
        $editedSite.PrimaryCircuit.SubnetMask = $EditControls.txtEditPrimarySubnetMask.Text.Trim()
        $editedSite.PrimaryCircuit.DefaultGateway = $EditControls.txtEditPrimaryDefaultGateway.Text.Trim()
        $editedSite.PrimaryCircuit.DNS1 = $EditControls.txtEditPrimaryDNS1.Text.Trim()
        $editedSite.PrimaryCircuit.DNS2 = $EditControls.txtEditPrimaryDNS2.Text.Trim()
        $editedSite.PrimaryCircuit.RouterModel = $EditControls.txtEditPrimaryRouterModel.Text.Trim()
        $editedSite.PrimaryCircuit.RouterName = $EditControls.txtEditPrimaryRouterName.Text.Trim()
        $editedSite.PrimaryCircuit.RouterSN = $EditControls.txtEditPrimaryRouterSN.Text.Trim()
        $editedSite.PrimaryCircuit.HasModem = $EditControls.chkEditPrimaryHasModem.IsChecked
        $editedSite.PrimaryCircuit.ModemModel = $EditControls.txtEditPrimaryModemModel.Text.Trim()
        $editedSite.PrimaryCircuit.ModemName = $EditControls.txtEditPrimaryModemName.Text.Trim()
        $editedSite.PrimaryCircuit.ModemSN = $EditControls.txtEditPrimaryModemSN.Text.Trim()
        
        # Backup Circuit
        $editedSite.HasBackupCircuit = $EditControls.chkEditHasBackupCircuit.IsChecked
        if ($editedSite.HasBackupCircuit) {
            $editedSite.BackupCircuit.Vendor = $EditControls.txtEditBackupVendor.Text.Trim()
            if ($EditControls.cmbEditBackupCircuitType.SelectedItem) {
                $editedSite.BackupCircuit.CircuitType = $EditControls.cmbEditBackupCircuitType.SelectedItem.Content
            }
            $editedSite.BackupCircuit.PPPoEUsername = $EditControls.txtEditBackupPPPoEUsername.Text.Trim()
            $editedSite.BackupCircuit.PPPoEPassword = $EditControls.txtEditBackupPPPoEPassword.Text.Trim()
            $editedSite.BackupCircuit.CircuitID = $EditControls.txtEditBackupCircuitID.Text.Trim()
            $editedSite.BackupCircuit.DownloadSpeed = $EditControls.txtEditBackupDownloadSpeed.Text.Trim()
            $editedSite.BackupCircuit.UploadSpeed = $EditControls.txtEditBackupUploadSpeed.Text.Trim()
            $editedSite.BackupCircuit.IPAddress = $EditControls.txtEditBackupIPAddress.Text.Trim()
            $editedSite.BackupCircuit.SubnetMask = $EditControls.txtEditBackupSubnetMask.Text.Trim()
            $editedSite.BackupCircuit.DefaultGateway = $EditControls.txtEditBackupDefaultGateway.Text.Trim()
            $editedSite.BackupCircuit.DNS1 = $EditControls.txtEditBackupDNS1.Text.Trim()
            $editedSite.BackupCircuit.DNS2 = $EditControls.txtEditBackupDNS2.Text.Trim()
            $editedSite.BackupCircuit.RouterModel = $EditControls.txtEditBackupRouterModel.Text.Trim()
            $editedSite.BackupCircuit.RouterName = $EditControls.txtEditBackupRouterName.Text.Trim()
            $editedSite.BackupCircuit.RouterSN = $EditControls.txtEditBackupRouterSN.Text.Trim()
            $editedSite.BackupCircuit.HasModem = $EditControls.chkEditBackupHasModem.IsChecked
            $editedSite.BackupCircuit.ModemModel = $EditControls.txtEditBackupModemModel.Text.Trim()
            $editedSite.BackupCircuit.ModemName = $EditControls.txtEditBackupModemName.Text.Trim()
            $editedSite.BackupCircuit.ModemSN = $EditControls.txtEditBackupModemSN.Text.Trim()
        }
        
        # VLANs
        $editedSite.VLANs.VLAN100_Servers = $EditControls.txtEditVlan100.Text.Trim()
        $editedSite.VLANs.VLAN101_NetworkDevices = $EditControls.txtEditVlan101.Text.Trim()
        $editedSite.VLANs.VLAN102_UserDevices = $EditControls.txtEditVlan102.Text.Trim()
        $editedSite.VLANs.VLAN103_UserDevices2 = $EditControls.txtEditVlan103.Text.Trim()
        $editedSite.VLANs.VLAN104_VOIP = $EditControls.txtEditVlan104.Text.Trim()
        $editedSite.VLANs.VLAN105_WiFiCorp = $EditControls.txtEditVlan105.Text.Trim()
        $editedSite.VLANs.VLAN106_WiFiBYOD = $EditControls.txtEditVlan106.Text.Trim()
        $editedSite.VLANs.VLAN107_WiFiGuest = $EditControls.txtEditVlan107.Text.Trim()
        $editedSite.VLANs.VLAN108_Spare = $EditControls.txtEditVlan108.Text.Trim()
        $editedSite.VLANs.VLAN109_DMZ = $EditControls.txtEditVlan109.Text.Trim()
        $editedSite.VLANs.VLAN110_CCTV = $EditControls.txtEditVlan110.Text.Trim()
        
        # Get device data from edit panels
        $editedSite.SwitchCount = if ($EditControls.cmbEditSwitchCount.SelectedItem) { [int]$EditControls.cmbEditSwitchCount.SelectedItem.Content } else { 1 }
        $editedSite.Switches = Get-EditDeviceDataFromUI 'Switch' $EditDeviceManager
        
        $editedSite.APCount = if ($EditControls.cmbEditAPCount.SelectedItem) { [int]$EditControls.cmbEditAPCount.SelectedItem.Content } else { 1 }
        $editedSite.AccessPoints = Get-EditDeviceDataFromUI 'AccessPoint' $EditDeviceManager
        
        $editedSite.UPSCount = if ($EditControls.cmbEditUPSCount.SelectedItem) { [int]$EditControls.cmbEditUPSCount.SelectedItem.Content } else { 0 }
        $editedSite.UPSDevices = Get-EditDeviceDataFromUI 'UPS' $EditDeviceManager
        
        $editedSite.CCTVCount = if ($EditControls.cmbEditCCTVCount.SelectedItem) { [int]$EditControls.cmbEditCCTVCount.SelectedItem.Content } else { 0 }
        $editedSite.CCTVDevices = Get-EditDeviceDataFromUI 'CCTV' $EditDeviceManager

        $editedSite.PrinterCount = if ($EditControls.cmbEditPrinterCount.SelectedItem) { [int]$EditControls.cmbEditPrinterCount.SelectedItem.Content } else { 0 }
        $editedSite.PrinterDevices = Get-EditDeviceDataFromUI 'Printer' $EditDeviceManager
        
        # Validate IPs
        [ValidationUtility]::ValidateDeviceIPs($editedSite)
        
        # Update the site in the data store
        if ($siteDataStore.UpdateEntry($editedSite)) {
            $EditControls.txtEditStatus.Text = "Site saved successfully!"
            $EditControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Green
            
            # Force a refresh of the DataGrid by re-setting the ItemsSource
            $dgSites.ItemsSource = $null
            $dgSites.ItemsSource = $siteDataStore.GetAllEntries()
            
            return $true
        } else {
            $EditControls.txtEditStatus.Text = "Failed to save site changes."
            $EditControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Red
            return $false
        }
        
    } catch {
        $EditControls.txtEditStatus.Text = "Error saving site: $($_.Exception.Message)"
        $EditControls.txtEditStatus.Foreground = [System.Windows.Media.Brushes]::Red
        return $false
    }
}

# Function to get device data from edit UI panels
function Get-EditDeviceDataFromUI {
    param(
        [string]$DeviceType,
        [object]$EditDeviceManager
    )
    
    $config = $EditDeviceManager.Configurations[$DeviceType]
    $stackPanel = $EditDeviceManager.StackPanels[$DeviceType]
    
    # Determine device type class name
    $className = switch ($DeviceType) {
        'Switch' { 'SwitchInfo' }
        'AccessPoint' { 'AccessPointInfo' }
        'UPS' { 'UPSInfo' }
        'CCTV' { 'CCTVInfo' }
        'Printer' { 'PrinterInfo' }
    }

    if ([string]::IsNullOrEmpty($className)) { 
    Write-Host "DEBUG: className is empty for DeviceType: $DeviceType"
}
    
    $devices = New-Object "System.Collections.Generic.List[$className]"
    Write-Host "DEBUG: Creating list for className: '$className', DeviceType: '$($Config.Type)'"
    # Get device count from the corresponding ComboBox
    $comboBoxName = "cmbEdit$DeviceType" + "Count"
    $comboBox = $EditDeviceManager.MainWindow.FindName($comboBoxName)
    $deviceCount = if ($comboBox.SelectedItem) { [int]$comboBox.SelectedItem.Content } else { 0 }
    
    for ($i = 1; $i -le $deviceCount; $i++) {
        $device = New-Object $className
        
        foreach ($groupBox in $stackPanel.Children) {
            if ($groupBox.Header -eq ($config.HeaderTemplate -f $i)) {
                foreach ($field in $config.Fields) {
                    $controlName = "txtEdit$DeviceType$i$field"
                    $control = $EditDeviceManager.FindControlInPanel($groupBox, $controlName)
                    if ($control) {
                        $device.$field = $control.Text.Trim()
                    }
                }
                break
            }
        }
        $devices.Add($device)
    }
    return $devices
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
        [System.Windows.MessageBox]::Show("Error initializing application: $_", "Initialization Error", "OK", "Error")
    }
})

# Show the window
    try {
        $mainWin.ShowDialog() | Out-Null
    } catch {
}