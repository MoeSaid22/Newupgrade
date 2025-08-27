# DeviceManager.ps1 - Device management classes and functions for Network Management

# Import required modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import data models
$dataModelsPath = Join-Path $scriptPath "DataModels.ps1"
if (Test-Path $dataModelsPath) {
    . $dataModelsPath
}

# Utility function - duplicated to avoid circular dependency
function Get-SafeValue {
    param([object]$Value)
    if ($Value) { return $Value.ToString() } else { return "" }
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
# DEVICE AND NETWORK AUTO-POPULATION FUNCTIONS
# ===================================================================

# Utility function to set ComboBox values safely
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
        
        # Auto-fill Site Subnet Code (VLAN identifier)
        if ($SiteSubnetCodeControl) {
            $SiteSubnetCodeControl.Text = $baseSubnet
        }
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
        
        # Auto-fill Site Subnet Code (VLAN identifier)
        if ($SiteSubnetCodeControl) {
            $SiteSubnetCodeControl.Text = $baseSubnet
        }
    }
}

