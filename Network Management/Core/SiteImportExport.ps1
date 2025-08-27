# SiteImportExport.ps1 - Import/Export functionality for Network Management

# Import data models
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataModelsPath = Join-Path $scriptPath "DataModels.ps1"
if (Test-Path $dataModelsPath) {
    . $dataModelsPath
} else {
    Write-Error "DataModels.ps1 not found at: $dataModelsPath"
}

# Helper function for safe COM object release
# COM Release Helper Function
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

# Helper function for safe values
function Get-SafeValue {
    param([object]$Value)
    if ($Value) { return $Value.ToString() } else { return "" }
}

# Import mode selection dialog
function Show-ImportModeDialog {
    # Create dialog with proper variable handling
    Add-Type -AssemblyName PresentationFramework
    
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Import Mode Selection" Height="350" Width="500"
    WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="How should duplicate sites be handled?" 
                   FontSize="14" FontWeight="Bold" Margin="0,0,0,20"/>
        
        <StackPanel Grid.Row="1" Margin="0,0,0,20">
            <RadioButton Name="rbUpdate" Content="Smart Update" 
                        GroupName="ImportMode" IsChecked="False" Margin="0,0,0,10"/>
            <TextBlock Text="Only update fields with new data`n   • Preserve existing data where Excel is empty" 
                      Foreground="Gray" Margin="0,0,0,15"/>
            
            <RadioButton Name="rbSkip" Content="Skip Duplicates (Recommended)" GroupName="ImportMode" 
                    IsChecked="True" Margin="0,0,0,10"/>
            <TextBlock Text="Existing sites will not be modified" 
                      Foreground="Gray" Margin="0,0,0,15"/>
            
            <RadioButton Name="rbReplace" Content="Replace Completely" 
                        GroupName="ImportMode" Margin="0,0,0,10"/>
            <TextBlock Text="WARNING: May lose existing data not in Excel" 
                      Foreground="DarkRed" Margin="0,0,0,15"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="Continue Import" Width="100" Height="30" 
                    Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="btnCancel" Content="Cancel" Width="75" Height="30" 
                    IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    $dialog.Owner = $mainWin
    
    # Get controls
    $rbSkip = $dialog.FindName("rbSkip")
    $rbUpdate = $dialog.FindName("rbUpdate") 
    $rbReplace = $dialog.FindName("rbReplace")
    $btnOK = $dialog.FindName("btnOK")
    $btnCancel = $dialog.FindName("btnCancel")
    
    # Set up event handlers with direct return values
    $btnOK.Add_Click({
        if ($rbUpdate.IsChecked) {
            $dialog.Tag = "Update"
        } elseif ($rbSkip.IsChecked) {
            $dialog.Tag = "Skip"
        } elseif ($rbReplace.IsChecked) {
            $dialog.Tag = "Replace"
        } else {
            $dialog.Tag = "Skip"  # Default to Skip
        }
        $dialog.DialogResult = $true
        $dialog.Close()
    })
    
    $btnCancel.Add_Click({
        $dialog.Tag = "Cancel"
        $dialog.DialogResult = $false
        $dialog.Close()
    })
    
    # Show dialog and return result
    $null = $dialog.ShowDialog()
    
    $result = $dialog.Tag
    if ([string]::IsNullOrEmpty($result)) {
        return "Cancel"
    }
    
    return $result
}

# Smart update function - only updates fields with new data
function Update-SiteWithNewData {
    param(
        [SiteEntry]$ExistingSite,
        [SiteEntry]$ImportSite
    )
    
    $changesCount = 0
    $changeDetails = @()
    
    # Pass both changes count and details by reference
    $changesRef = [ref]$changesCount
    $changeDetailsRef = [ref]$changeDetails
    
    # Helper function to update field only if import has data AND it's different
function Update-FieldIfNotEmpty {
    param($ExistingObject, $ImportObject, $PropertyName, [ref]$ChangesRef, [ref]$ChangeDetailsRef)
    
    $importValue = $ImportObject.$PropertyName
    $existingValue = $ExistingObject.$PropertyName
    
    if (-not [string]::IsNullOrWhiteSpace($importValue)) {
        # Convert both to strings for comparison and trim whitespace
        $importStr = $importValue.ToString().Trim()
        $existingStr = if ($existingValue) { $existingValue.ToString().Trim() } else { "" }
        
        if ($existingStr -ne $importStr) {
            $ExistingObject.$PropertyName = $importValue
            $ChangesRef.Value++
            
            # Add human-readable change description for ALL possible fields
            $fieldDisplayName = switch ($PropertyName) {
                # Basic Info Fields
                "SiteCode" { "Site Code" }
                "SiteSubnet" { "Site Subnet" }
                "SiteSubnetCode" { "Site Subnet Code" }
                "SiteName" { "Site Name" }
                "SiteAddress" { "Site Address" }
                "MainContactName" { "Main Contact Name" }
                "MainContactPhone" { "Main Contact Phone" }
                "SecondContactName" { "Second Contact Name" }
                "SecondContactPhone" { "Second Contact Phone" }
                
                # Firewall Fields
                "FirewallIP" { "Firewall IP" }
                "FirewallName" { "Firewall Name" }
                "FirewallVersion" { "Firewall Version" }
                "FirewallSN" { "Firewall Serial Number" }
                
                # Circuit Fields
                "Vendor" { "Vendor" }
                "CircuitType" { "Circuit Type" }
                "CircuitID" { "Circuit ID" }
                "DownloadSpeed" { "Download Speed" }
                "UploadSpeed" { "Upload Speed" }
                "IPAddress" { "IP Address" }
                "SubnetMask" { "Subnet Mask" }
                "DefaultGateway" { "Default Gateway" }
                "DNS1" { "Primary DNS" }
                "DNS2" { "Secondary DNS" }
                "RouterModel" { "Router Model" }
                "RouterName" { "Router Name" }
                "RouterSN" { "Router Serial Number" }
                "PPPoEUsername" { "PPPoE Username" }
                "PPPoEPassword" { "PPPoE Password" }
                "ModemModel" { "Modem Model" }
                "ModemName" { "Modem Name" }
                "ModemSN" { "Modem Serial Number" }
                
                # Device Fields
                "ManagementIP" { "Management IP" }
                "Name" { "Device Name" }
                "AssetTag" { "Asset Tag" }
                "Version" { "Version" }
                "SerialNumber" { "Serial Number" }
                
                # VLAN Fields
                "VLAN100_Servers" { "VLAN 100 (Servers)" }
                "VLAN101_NetworkDevices" { "VLAN 101 (Network Devices)" }
                "VLAN102_UserDevices" { "VLAN 102 (User Devices)" }
                "VLAN103_UserDevices2" { "VLAN 103 (User Devices 2)" }
                "VLAN104_VOIP" { "VLAN 104 (VOIP)" }
                "VLAN105_WiFiCorp" { "VLAN 105 (WiFi Corporate)" }
                "VLAN106_WiFiBYOD" { "VLAN 106 (WiFi BYOD)" }
                "VLAN107_WiFiGuest" { "VLAN 107 (WiFi Guest)" }
                "VLAN108_Spare" { "VLAN 108 (Spare)" }
                "VLAN109_DMZ" { "VLAN 109 (DMZ)" }
                "VLAN110_CCTV" { "VLAN 110 (CCTV)" }
                
                default { $PropertyName }
            }
            
            $ChangeDetailsRef.Value += $fieldDisplayName
            return $true
        }
    }
    return $false
}
    
    # Pass changes count by reference to all calls
    $changesRef = [ref]$changesCount
    
    # Update basic info
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteSubnet" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteSubnetCode" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SiteAddress" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "MainContactName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "MainContactPhone" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SecondContactName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "SecondContactPhone" $changesRef $changeDetailsRef

    # Update firewall info
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallIP" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallVersion" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite $ImportSite "FirewallSN" $changesRef $changeDetailsRef
    
    # Update device counts if import has higher counts
    if ($ImportSite.SwitchCount -gt $ExistingSite.SwitchCount) {
        $ExistingSite.SwitchCount = $ImportSite.SwitchCount
        $changesCount++
    }
    if ($ImportSite.APCount -gt $ExistingSite.APCount) {
        $ExistingSite.APCount = $ImportSite.APCount
        $changesCount++
    }
    if ($ImportSite.UPSCount -gt $ExistingSite.UPSCount) {
        $ExistingSite.UPSCount = $ImportSite.UPSCount
        $changesCount++
    }
    if ($ImportSite.CCTVCount -gt $ExistingSite.CCTVCount) {
        $ExistingSite.CCTVCount = $ImportSite.CCTVCount
        $changesCount++
    }
    
    # Update devices - expand arrays if needed and update individual devices
    # Switches
    while ($ExistingSite.Switches.Count -lt $ImportSite.Switches.Count) {
        $ExistingSite.Switches.Add([SwitchInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.Switches.Count; $i++) {
        if ($i -lt $ExistingSite.Switches.Count) {
            $deviceNum = $i + 1
            
            # Check each field and add device number to change description
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                # Replace the last added item with device-specific description
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "AssetTag" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Asset Tag"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "Version" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Version"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.Switches[$i] $ImportSite.Switches[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Switch $deviceNum Serial Number"
            }
        }
    }

    # Access Points
    while ($ExistingSite.AccessPoints.Count -lt $ImportSite.AccessPoints.Count) {
        $ExistingSite.AccessPoints.Add([AccessPointInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.AccessPoints.Count; $i++) {
        if ($i -lt $ExistingSite.AccessPoints.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "AssetTag" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Asset Tag"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "Version" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Version"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.AccessPoints[$i] $ImportSite.AccessPoints[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "AP $deviceNum Serial Number"
            }
        }
    }

    # UPS Devices
    while ($ExistingSite.UPSDevices.Count -lt $ImportSite.UPSDevices.Count) {
        $ExistingSite.UPSDevices.Add([UPSInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.UPSDevices.Count; $i++) {
        if ($i -lt $ExistingSite.UPSDevices.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.UPSDevices[$i] $ImportSite.UPSDevices[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "UPS $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.UPSDevices[$i] $ImportSite.UPSDevices[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "UPS $deviceNum Name"
            }
        }
    }

    # CCTV Devices
    while ($ExistingSite.CCTVDevices.Count -lt $ImportSite.CCTVDevices.Count) {
        $ExistingSite.CCTVDevices.Add([CCTVInfo]::new())
    }
    for ($i = 0; $i -lt $ImportSite.CCTVDevices.Count; $i++) {
        if ($i -lt $ExistingSite.CCTVDevices.Count) {
            $deviceNum = $i + 1
            
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "ManagementIP" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Management IP"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "Name" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Name"
            }
            if (Update-FieldIfNotEmpty $ExistingSite.CCTVDevices[$i] $ImportSite.CCTVDevices[$i] "SerialNumber" $changesRef $changeDetailsRef) {
                $changeDetails[$changeDetails.Count - 1] = "Camera $deviceNum Serial Number"
            }
        }
    }

    # Update ALL circuit fields
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "Vendor" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "CircuitType" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "CircuitID" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DownloadSpeed" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "UploadSpeed" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "IPAddress" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "SubnetMask" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DefaultGateway" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DNS1" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "DNS2" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterModel" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "RouterSN" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "PPPoEUsername" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "PPPoEPassword" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemModel" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemName" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.PrimaryCircuit $ImportSite.PrimaryCircuit "ModemSN" $changesRef $changeDetailsRef

    # Update backup circuit if import has backup data
    if ($ImportSite.HasBackupCircuit) {
        if (-not $ExistingSite.HasBackupCircuit) {
            $ExistingSite.HasBackupCircuit = $true
            $changesCount++
            $changeDetails += "Added Backup Circuit"
        }
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "Vendor" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "CircuitType" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "CircuitID" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DownloadSpeed" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "UploadSpeed" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "IPAddress" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "SubnetMask" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DefaultGateway" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DNS1" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "DNS2" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterModel" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterName" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "RouterSN" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "PPPoEUsername" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "PPPoEPassword" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemModel" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemName" $changesRef $changeDetailsRef
        Update-FieldIfNotEmpty $ExistingSite.BackupCircuit $ImportSite.BackupCircuit "ModemSN" $changesRef $changeDetailsRef
    }

    # Update ALL VLAN fields
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN100_Servers" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN101_NetworkDevices" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN102_UserDevices" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN103_UserDevices2" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN104_VOIP" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN105_WiFiCorp" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN106_WiFiBYOD" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN107_WiFiGuest" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN108_Spare" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN109_DMZ" $changesRef $changeDetailsRef
    Update-FieldIfNotEmpty $ExistingSite.VLANs $ImportSite.VLANs "VLAN110_CCTV" $changesRef $changeDetailsRef
    
    # Update display properties
    $ExistingSite.UpdateDisplayProperties()
    
    # Return both the site and whether changes were made
    return @{
        Site = $ExistingSite
        HasChanges = ($changesCount -gt 0)
        ChangesCount = $changesCount
        ChangeDetails = $changeDetails
    }
}

# Export sites to Excel with 9 sheets matching sub-tabs exactly
function Export-SitesToExcel {
    param([string]$FilePath)
   
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        throw "File path cannot be empty"
    }

    # Initialize COM objects to null
        $excel = $null
        $workbook = $null
    
    try {
        $allSites = $siteDataStore.GetAllEntries()
        
        # Show progress panel (same as import)
        $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Visible
        $pbSiteImportProgress.Value = 0
        $txtSiteProgressStatus.Text = "Starting Excel export..."
        $txtSiteProgressDetails.Text = ""
        
        # Try to create Excel COM object
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $txtSiteProgressStatus.Text = "Excel application started, creating workbook..."
            $pbSiteImportProgress.Value = 5
            [System.Windows.Forms.Application]::DoEvents()
        }
        catch {
            throw "Excel is not installed or cannot be accessed. Please install Microsoft Excel to use this feature."
        }
        
        try {
            $workbook = $excel.Workbooks.Add()
            
            # Helper function to write data efficiently
            function Write-SheetData {
                param($Sheet, $Headers, $DataRows)
                
                # Write headers
                for ($col = 0; $col -lt $Headers.Count; $col++) {
                    $Sheet.Cells.Item(1, $col + 1).Value2 = $Headers[$col]
                    $Sheet.Cells.Item(1, $col + 1).Font.Bold = $true
                }
                
                # Write data rows
                for ($row = 0; $row -lt $DataRows.Count; $row++) {
                    $dataRow = $DataRows[$row]
                    for ($col = 0; $col -lt $dataRow.Count; $col++) {
                        $Sheet.Cells.Item($row + 2, $col + 1).Value2 = $dataRow[$col]
                    }
                }
                
                $Sheet.Columns.AutoFit() | Out-Null
            }
            
            $txtSiteProgressStatus.Text = "Creating worksheets..."
            $pbSiteImportProgress.Value = 10
            [System.Windows.Forms.Application]::DoEvents()
            
            # Delete all default sheets first
            while ($workbook.Worksheets.Count -gt 1) {
                $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
            }
            
            # Rename the remaining sheet to Basic Info
            $basicSheet = $workbook.Worksheets.Item(1)
            $basicSheet.Name = "Basic Info"
            
            # Create remaining sheets in correct order
            $switchSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $basicSheet)
            $switchSheet.Name = "Switches"
            
            $apSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $switchSheet)
            $apSheet.Name = "Access Points"
            
            $firewallSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $apSheet)
            $firewallSheet.Name = "Firewall"
            
            $primarySheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $firewallSheet)
            $primarySheet.Name = "Primary Circuit"
            
            $backupSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $primarySheet)
            $backupSheet.Name = "Backup Circuit"
            
            $vlanSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $backupSheet)
            $vlanSheet.Name = "VLANs"
            
            $cctvSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $vlanSheet)
            $cctvSheet.Name = "CCTV"
            
            $upsSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $cctvSheet)
            $upsSheet.Name = "UPS"

            $printerSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $upsSheet)
            $printerSheet.Name = "Printer"
            
            # ============================================================================
            # 1. BASIC INFO SHEET
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Basic Info..."
            $txtSiteProgressDetails.Text = "Processing basic site information..."
            $pbSiteImportProgress.Value = 15
            [System.Windows.Forms.Application]::DoEvents()
            
            $basicHeaders = @("SiteID", "SiteCode", "SiteSubnet", "SiteSubnetCode", "SiteName", "SiteAddress", "MainContactName", "MainContactPhone", "SecondContactName", "SecondContactPhone")
            
            $basicDataRows = @()
            foreach ($site in $allSites) {
                $basicDataRows += ,@(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.SiteSubnet),
                    (Get-SafeValue $site.SiteSubnetCode),
                    (Get-SafeValue $site.SiteName),
                    (Get-SafeValue $site.SiteAddress),
                    (Get-SafeValue $site.MainContactName),
                    (Get-SafeValue $site.MainContactPhone),
                    (Get-SafeValue $site.SecondContactName),
                    (Get-SafeValue $site.SecondContactPhone)
                )
            }
            
            Write-SheetData -Sheet $basicSheet -Headers $basicHeaders -DataRows $basicDataRows
            
            # ============================================================================
            # 2. SWITCHES SHEET (10 switches max)
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Switches..."
            $txtSiteProgressDetails.Text = "Processing switch data for $($allSites.Count) sites..."
            $pbSiteImportProgress.Value = 25
            [System.Windows.Forms.Application]::DoEvents()
            
            $switchHeaders = @("SiteID", "SiteCode", "SwitchCount")
            for ($i = 1; $i -le 10; $i++) {
                $switchHeaders += "SW${i}_ManagementIP"
                $switchHeaders += "SW${i}_Name"
                $switchHeaders += "SW${i}_AssetTag"
                $switchHeaders += "SW${i}_Version"
                $switchHeaders += "SW${i}_SerialNumber"
            }
            
            $switchDataRows = @()
            foreach ($site in $allSites) {
                $row = @(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.SwitchCount)
                )
                
                for ($i = 0; $i -lt 10; $i++) {
                    if ($i -lt $site.Switches.Count) {
                        $switch = $site.Switches[$i]
                        $row += (Get-SafeValue $switch.ManagementIP)
                        $row += (Get-SafeValue $switch.Name)
                        $row += (Get-SafeValue $switch.AssetTag)
                        $row += (Get-SafeValue $switch.Version)
                        $row += (Get-SafeValue $switch.SerialNumber)
                    } else {
                        $row += "", "", "", "", ""
                    }
                }
                $switchDataRows += ,$row
            }
            
            Write-SheetData -Sheet $switchSheet -Headers $switchHeaders -DataRows $switchDataRows
            
            # ============================================================================
            # 3. ACCESS POINTS SHEET (10 APs max)
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Access Points..."
            $txtSiteProgressDetails.Text = "Processing access point data..."
            $pbSiteImportProgress.Value = 35
            [System.Windows.Forms.Application]::DoEvents()
            
            $apHeaders = @("SiteID", "SiteCode", "APCount")
            for ($i = 1; $i -le 10; $i++) {
                $apHeaders += "AP${i}_ManagementIP"
                $apHeaders += "AP${i}_Name"
                $apHeaders += "AP${i}_AssetTag"
                $apHeaders += "AP${i}_Version"
                $apHeaders += "AP${i}_SerialNumber"
            }
            
            $apDataRows = @()
            foreach ($site in $allSites) {
                $row = @(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.APCount)
                )
                
                for ($i = 0; $i -lt 10; $i++) {
                    if ($i -lt $site.AccessPoints.Count) {
                        $ap = $site.AccessPoints[$i]
                        $row += (Get-SafeValue $ap.ManagementIP)
                        $row += (Get-SafeValue $ap.Name)
                        $row += (Get-SafeValue $ap.AssetTag)
                        $row += (Get-SafeValue $ap.Version)
                        $row += (Get-SafeValue $ap.SerialNumber)
                    } else {
                        $row += "", "", "", "", ""
                    }
                }
                $apDataRows += ,$row
            }
            
            Write-SheetData -Sheet $apSheet -Headers $apHeaders -DataRows $apDataRows
            
            # ============================================================================
            # 4. FIREWALL SHEET
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Firewall..."
            $txtSiteProgressDetails.Text = "Processing firewall configuration data..."
            $pbSiteImportProgress.Value = 45
            [System.Windows.Forms.Application]::DoEvents()
            
            $firewallHeaders = @("SiteID", "SiteCode", "FirewallIP", "FirewallName", "FirewallVersion", "FirewallSN")
            
            $firewallDataRows = @()
            foreach ($site in $allSites) {
                $firewallDataRows += ,@(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.FirewallIP),
                    (Get-SafeValue $site.FirewallName),
                    (Get-SafeValue $site.FirewallVersion),
                    (Get-SafeValue $site.FirewallSN)
                )
            }
            
            Write-SheetData -Sheet $firewallSheet -Headers $firewallHeaders -DataRows $firewallDataRows
            
            # ============================================================================
            # 5. PRIMARY CIRCUIT SHEET
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Primary Circuit..."
            $txtSiteProgressDetails.Text = "Processing primary circuit configuration..."
            $pbSiteImportProgress.Value = 55
            [System.Windows.Forms.Application]::DoEvents()
            
            $primaryHeaders = @("SiteID", "SiteCode", "Vendor", "CircuitType", "CircuitID", "DownloadSpeed", "UploadSpeed", "IPAddress", "SubnetMask", "DefaultGateway", "DNS1", "DNS2", "RouterModel", "RouterName", "RouterSN", "PPPoEUsername", "PPPoEPassword", "HasModem", "ModemModel", "ModemName", "ModemSN")
            
            $primaryDataRows = @()
            foreach ($site in $allSites) {
                $primaryDataRows += ,@(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.PrimaryCircuit.Vendor),
                    (Get-SafeValue $site.PrimaryCircuit.CircuitType),
                    (Get-SafeValue $site.PrimaryCircuit.CircuitID),
                    (Get-SafeValue $site.PrimaryCircuit.DownloadSpeed),
                    (Get-SafeValue $site.PrimaryCircuit.UploadSpeed),
                    (Get-SafeValue $site.PrimaryCircuit.IPAddress),
                    (Get-SafeValue $site.PrimaryCircuit.SubnetMask),
                    (Get-SafeValue $site.PrimaryCircuit.DefaultGateway),
                    (Get-SafeValue $site.PrimaryCircuit.DNS1),
                    (Get-SafeValue $site.PrimaryCircuit.DNS2),
                    (Get-SafeValue $site.PrimaryCircuit.RouterModel),
                    (Get-SafeValue $site.PrimaryCircuit.RouterName),
                    (Get-SafeValue $site.PrimaryCircuit.RouterSN),
                    (Get-SafeValue $site.PrimaryCircuit.PPPoEUsername),
                    (Get-SafeValue $site.PrimaryCircuit.PPPoEPassword),
                    (Get-SafeValue $site.PrimaryCircuit.HasModem),
                    (Get-SafeValue $site.PrimaryCircuit.ModemModel),
                    (Get-SafeValue $site.PrimaryCircuit.ModemName),
                    (Get-SafeValue $site.PrimaryCircuit.ModemSN)
                )
            }
            
            Write-SheetData -Sheet $primarySheet -Headers $primaryHeaders -DataRows $primaryDataRows
            
            # ============================================================================
            # 6. BACKUP CIRCUIT SHEET
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Backup Circuit..."
            $txtSiteProgressDetails.Text = "Processing backup circuit configuration..."
            $pbSiteImportProgress.Value = 65
            [System.Windows.Forms.Application]::DoEvents()
            
            $backupHeaders = @("SiteID", "SiteCode", "HasBackupCircuit", "Vendor", "CircuitType", "CircuitID", "DownloadSpeed", "UploadSpeed", "IPAddress", "SubnetMask", "DefaultGateway", "DNS1", "DNS2", "RouterModel", "RouterName", "RouterSN", "PPPoEUsername", "PPPoEPassword", "HasModem", "ModemModel", "ModemName", "ModemSN")
            
            $backupDataRows = @()
            foreach ($site in $allSites) {
                $backupDataRows += ,@(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.HasBackupCircuit),
                    (Get-SafeValue $site.BackupCircuit.Vendor),
                    (Get-SafeValue $site.BackupCircuit.CircuitType),
                    (Get-SafeValue $site.BackupCircuit.CircuitID),
                    (Get-SafeValue $site.BackupCircuit.DownloadSpeed),
                    (Get-SafeValue $site.BackupCircuit.UploadSpeed),
                    (Get-SafeValue $site.BackupCircuit.IPAddress),
                    (Get-SafeValue $site.BackupCircuit.SubnetMask),
                    (Get-SafeValue $site.BackupCircuit.DefaultGateway),
                    (Get-SafeValue $site.BackupCircuit.DNS1),
                    (Get-SafeValue $site.BackupCircuit.DNS2),
                    (Get-SafeValue $site.BackupCircuit.RouterModel),
                    (Get-SafeValue $site.BackupCircuit.RouterName),
                    (Get-SafeValue $site.BackupCircuit.RouterSN),
                    (Get-SafeValue $site.BackupCircuit.PPPoEUsername),
                    (Get-SafeValue $site.BackupCircuit.PPPoEPassword),
                    (Get-SafeValue $site.BackupCircuit.HasModem),
                    (Get-SafeValue $site.BackupCircuit.ModemModel),
                    (Get-SafeValue $site.BackupCircuit.ModemName),
                    (Get-SafeValue $site.BackupCircuit.ModemSN)
                )
            }
            
            Write-SheetData -Sheet $backupSheet -Headers $backupHeaders -DataRows $backupDataRows
            
            # ============================================================================
            # 7. VLANS SHEET
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting VLANs..."
            $txtSiteProgressDetails.Text = "Processing VLAN configuration..."
            $pbSiteImportProgress.Value = 75
            [System.Windows.Forms.Application]::DoEvents()
            
            $vlanHeaders = @("SiteID", "SiteCode", "VLAN100_Servers", "VLAN101_NetworkDevices", "VLAN102_UserDevices", "VLAN103_UserDevices2", "VLAN104_VOIP", "VLAN105_WiFiCorp", "VLAN106_WiFiBYOD", "VLAN107_WiFiGuest", "VLAN108_Spare", "VLAN109_DMZ", "VLAN110_CCTV")
            
            $vlanDataRows = @()
            foreach ($site in $allSites) {
                $vlanDataRows += ,@(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.VLANs.VLAN100_Servers),
                    (Get-SafeValue $site.VLANs.VLAN101_NetworkDevices),
                    (Get-SafeValue $site.VLANs.VLAN102_UserDevices),
                    (Get-SafeValue $site.VLANs.VLAN103_UserDevices2),
                    (Get-SafeValue $site.VLANs.VLAN104_VOIP),
                    (Get-SafeValue $site.VLANs.VLAN105_WiFiCorp),
                    (Get-SafeValue $site.VLANs.VLAN106_WiFiBYOD),
                    (Get-SafeValue $site.VLANs.VLAN107_WiFiGuest),
                    (Get-SafeValue $site.VLANs.VLAN108_Spare),
                    (Get-SafeValue $site.VLANs.VLAN109_DMZ),
                    (Get-SafeValue $site.VLANs.VLAN110_CCTV)
                )
            }
            
            Write-SheetData -Sheet $vlanSheet -Headers $vlanHeaders -DataRows $vlanDataRows
            
            # ============================================================================
            # 8. CCTV SHEET (10 cameras max)
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting CCTV..."
            $txtSiteProgressDetails.Text = "Processing CCTV camera data..."
            $pbSiteImportProgress.Value = 85
            [System.Windows.Forms.Application]::DoEvents()
            
            $cctvHeaders = @("SiteID", "SiteCode", "CCTVCount")
            for ($i = 1; $i -le 10; $i++) {
                $cctvHeaders += "CAM${i}_ManagementIP"
                $cctvHeaders += "CAM${i}_Name"
                $cctvHeaders += "CAM${i}_SerialNumber"
            }
            
            $cctvDataRows = @()
            foreach ($site in $allSites) {
                $row = @(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.CCTVCount)
                )
                
                for ($i = 0; $i -lt 10; $i++) {
                    if ($i -lt $site.CCTVDevices.Count) {
                        $cctv = $site.CCTVDevices[$i]
                        $row += (Get-SafeValue $cctv.ManagementIP)
                        $row += (Get-SafeValue $cctv.Name)
                        $row += (Get-SafeValue $cctv.SerialNumber)
                    } else {
                        $row += "", "", ""
                    }
                }
                $cctvDataRows += ,$row
            }
            
            Write-SheetData -Sheet $cctvSheet -Headers $cctvHeaders -DataRows $cctvDataRows
            
            # ============================================================================
            # 9. UPS SHEET (5 UPS max)
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting UPS..."
            $txtSiteProgressDetails.Text = "Processing UPS device data..."
            $pbSiteImportProgress.Value = 95
            [System.Windows.Forms.Application]::DoEvents()
            
            $upsHeaders = @("SiteID", "SiteCode", "UPSCount")
            for ($i = 1; $i -le 5; $i++) {
                $upsHeaders += "UPS${i}_ManagementIP"
                $upsHeaders += "UPS${i}_Name"
            }
            
            $upsDataRows = @()
            foreach ($site in $allSites) {
                $row = @(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.UPSCount)
                )
                
                for ($i = 0; $i -lt 5; $i++) {
                    if ($i -lt $site.UPSDevices.Count) {
                        $ups = $site.UPSDevices[$i]
                        $row += (Get-SafeValue $ups.ManagementIP)
                        $row += (Get-SafeValue $ups.Name)
                    } else {
                        $row += "", ""
                    }
                }
                $upsDataRows += ,$row
            }
            
            Write-SheetData -Sheet $upsSheet -Headers $upsHeaders -DataRows $upsDataRows
            
            # ============================================================================
            # 10. PRINTER SHEET (10 printers max)
            # ============================================================================
            $txtSiteProgressStatus.Text = "Exporting Printers..."
            $txtSiteProgressDetails.Text = "Processing printer device data..."
            $pbSiteImportProgress.Value = 98
            [System.Windows.Forms.Application]::DoEvents()

            $printerHeaders = @("SiteID", "SiteCode", "PrinterCount")
            for ($i = 1; $i -le 10; $i++) {
                $printerHeaders += "PRT${i}_ManagementIP"
                $printerHeaders += "PRT${i}_Name"
                $printerHeaders += "PRT${i}_Model"
                $printerHeaders += "PRT${i}_SerialNumber"
            }

            $printerDataRows = @()
            foreach ($site in $allSites) {
                $row = @(
                    (Get-SafeValue $site.ID),
                    (Get-SafeValue $site.SiteCode),
                    (Get-SafeValue $site.PrinterCount)
                )
                
                for ($i = 0; $i -lt 10; $i++) {
                    if ($i -lt $site.PrinterDevices.Count) {
                        $printer = $site.PrinterDevices[$i]
                        $row += (Get-SafeValue $printer.ManagementIP)
                        $row += (Get-SafeValue $printer.Name)
                        $row += (Get-SafeValue $printer.Model)
                        $row += (Get-SafeValue $printer.SerialNumber)
                    } else {
                        $row += "", "", "", ""
                    }
                }
                $printerDataRows += ,$row
            }

            Write-SheetData -Sheet $printerSheet -Headers $printerHeaders -DataRows $printerDataRows
            # Save and close
            $txtSiteProgressStatus.Text = "Saving Excel file..."
            $txtSiteProgressDetails.Text = "Finalizing export to: $([System.IO.Path]::GetFileName($FilePath))"
            $pbSiteImportProgress.Value = 98
            [System.Windows.Forms.Application]::DoEvents()
            
            $workbook.SaveAs($FilePath)
            $workbook.Close()
            
            $pbSiteImportProgress.Value = 100
            $txtSiteProgressStatus.Text = "Export completed successfully!"
            $txtSiteProgressDetails.Text = "Exported $($allSites.Count) sites to 10 worksheets"
            [System.Windows.Forms.Application]::DoEvents()
            
            return "Successfully exported $($allSites.Count) sites to Excel file with 10 separate worksheets:`n$FilePath"
        }
        finally {
            if ($workbook) { 
                try { $workbook.Close() } catch { }
            }
        }
    }
    catch {
        throw $_.Exception.Message
    }
    
        finally {
        # CRITICAL: Enhanced cleanup to eliminate all Excel processes
        
        # Close and release workbook
        if ($workbook) { 
            try { 
                $workbook.Close($false)  # Close without saving changes
            } catch { }
            $workbook = Release-ComObject $workbook
        }
        
        # Quit and release Excel application
        if ($excel) { 
            try { 
                $excel.DisplayAlerts = $true
                $excel.Quit()
            } catch { }
            $excel = Release-ComObject $excel
        }
        
        # Multiple rounds of garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Small delay to ensure process termination
        Start-Sleep -Milliseconds 500
        
        # Hide progress panel when done
        $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Import sites from Excel - Robust round-trip compatible with full validation
function Import-SitesFromExcel {
    param([string]$ExcelFilePath)
    
    if ([string]::IsNullOrWhiteSpace($ExcelFilePath)) {
        throw "File path cannot be empty"
    }
   
    if (-not (Test-Path $ExcelFilePath)) {
        throw "Excel file not found"
    }
   
    if (-not ($ExcelFilePath.ToLower().EndsWith('.xlsx') -or $ExcelFilePath.ToLower().EndsWith('.xls'))) {
        throw "File must be an Excel file (.xlsx or .xls extension)"
    }
   
    $fileInfo = Get-Item $ExcelFilePath
    if ($fileInfo.Length -gt 100MB) {
        throw "File is too large. Maximum size is 100MB"
    }

    $script:SubnetIssueDetails = @{}
    
    # Initialize COM objects to null
    $excel = $null
    $workbook = $null
    
    try {
        # Show progress panel
        $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Visible
        $pbSiteImportProgress.Value = 0
        $txtSiteProgressStatus.Text = "Starting Excel import..."
        $txtSiteProgressDetails.Text = ""
        
        # Try to create Excel COM object
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
        }
        catch {
            throw "Excel is not installed or cannot be accessed. Please install Microsoft Excel to use this feature."
        }
        
        # Update progress for file opening
        $txtSiteProgressStatus.Text = "Opening Excel file..."
        $txtSiteProgressDetails.Text = "Loading: $([System.IO.Path]::GetFileName($ExcelFilePath))"
        $pbSiteImportProgress.Value = 3
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $workbook = $excel.Workbooks.Open($ExcelFilePath)
        }
        catch {
            throw "Failed to open Excel file: $_"
        }
        
        # Confirm file opened
        $txtSiteProgressStatus.Text = "Excel file opened successfully"
        $pbSiteImportProgress.Value = 6
        [System.Windows.Forms.Application]::DoEvents()
        
        # Helper function to safely get cell value
        function Get-CellValue {
            param($Sheet, $Row, $Col)
            try {
                $value = $Sheet.Cells.Item($Row, $Col).Value2
                if ($null -eq $value) { return "" }
                return $value.ToString().Trim()
            }
            catch {
                return ""
            }
        }
        
        # Helper function to find sheet by name
        function Get-SheetByName {
            param($Workbook, $SheetName)
            try {
                return $Workbook.Worksheets.Item($SheetName)
            }
            catch {
                return $null
            }
        }
        
        # Check if required sheets exist (all 9 sheets)
        $requiredSheets = @("Basic Info", "Switches", "Access Points", "Firewall", "Primary Circuit", "Backup Circuit", "VLANs", "CCTV", "UPS", "Printer")
        $missingSheets = @()
        
        foreach ($sheetName in $requiredSheets) {
            $sheet = Get-SheetByName -Workbook $workbook -SheetName $sheetName
            if ($null -eq $sheet) {
                $missingSheets += $sheetName
            }
        }
        
        if ($missingSheets.Count -gt 0) {
            throw "Missing required sheets: $($missingSheets -join ', '). This doesn't appear to be a valid export file."
        }
        
        # Get all sheets
        $basicSheet = Get-SheetByName -Workbook $workbook -SheetName "Basic Info"
        $switchSheet = Get-SheetByName -Workbook $workbook -SheetName "Switches"
        $apSheet = Get-SheetByName -Workbook $workbook -SheetName "Access Points"
        $firewallSheet = Get-SheetByName -Workbook $workbook -SheetName "Firewall"
        $primarySheet = Get-SheetByName -Workbook $workbook -SheetName "Primary Circuit"
        $backupSheet = Get-SheetByName -Workbook $workbook -SheetName "Backup Circuit"
        $vlanSheet = Get-SheetByName -Workbook $workbook -SheetName "VLANs"
        $cctvSheet = Get-SheetByName -Workbook $workbook -SheetName "CCTV"
        $upsSheet = Get-SheetByName -Workbook $workbook -SheetName "UPS"
        $printerSheet = Get-SheetByName -Workbook $workbook -SheetName "Printer"
        
        # Find the last row with data in Basic Info sheet
        $lastRow = 1
        for ($row = 2; $row -le 1000; $row++) {
            $siteCode = Get-CellValue -Sheet $basicSheet -Row $row -Col 2  # SiteCode column
            if ([string]::IsNullOrWhiteSpace($siteCode)) {
                break
            }
            $lastRow = $row
        }
        
        if ($lastRow -eq 1) {
            throw "No data found in the Excel file"
        }
        
        $totalSites = $lastRow - 1
        $importedCount = 0
        $skippedCount = 0
        $errorMessages = @()

        # Site tracking for detailed reporting
        $script:ProcessedSites = @()

        $skippedDuplicates = @()
        $errorSites = @()
        
        # Get import mode from user - MOVED HERE TO ASK ONLY ONCE
        $importMode = Show-ImportModeDialog
        if ($importMode -eq "Cancel") {
            return "Import cancelled by user"
        }

        # DEBUG: Validate the import mode
        if ($importMode -notin @("Skip", "Update", "Replace")) {
            throw "Invalid import mode received: '$importMode'. Expected Skip, Update, or Replace."
        }

        $txtSiteProgressStatus.Text = "Import mode: $importMode - Processing $totalSites sites..."
        $txtSiteProgressDetails.Text = "Processing with $importMode mode for duplicates..."
        $pbSiteImportProgress.Value = 10
        [System.Windows.Forms.Application]::DoEvents()
        
        # Process each site (rows 2 onwards, row 1 is headers)
        for ($row = 2; $row -le $lastRow; $row++) {
            try {
                # Create new site entry
                $site = [SiteEntry]::new()
                
                # BASIC INFO DATA
                $site.SiteCode = Get-CellValue -Sheet $basicSheet -Row $row -Col 2
                $site.SiteSubnet = Get-CellValue -Sheet $basicSheet -Row $row -Col 3
                $site.SiteSubnetCode = Get-CellValue -Sheet $basicSheet -Row $row -Col 4
                $site.SiteName = Get-CellValue -Sheet $basicSheet -Row $row -Col 5
                $site.SiteAddress = Get-CellValue -Sheet $basicSheet -Row $row -Col 6
                $site.MainContactName = Get-CellValue -Sheet $basicSheet -Row $row -Col 7
                $site.MainContactPhone = Get-CellValue -Sheet $basicSheet -Row $row -Col 8
                $site.SecondContactName = Get-CellValue -Sheet $basicSheet -Row $row -Col 9
                $site.SecondContactPhone = Get-CellValue -Sheet $basicSheet -Row $row -Col 10
                
                # Validate required fields
                if ([string]::IsNullOrWhiteSpace($site.SiteCode) -or [string]::IsNullOrWhiteSpace($site.SiteName)) {
                    throw "Missing required fields (SiteCode and SiteName are mandatory)"
                }
                
                # VALIDATE SUBNET AND TRACK ISSUES (but don't throw errors)
                $subnetIssues = @()
                $shouldUpdateSubnet = $true
                if (-not [string]::IsNullOrWhiteSpace($site.SiteSubnet)) {   
                    # Validate subnet format
                    if ($site.SiteSubnet -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                        $octets = $site.SiteSubnet.Split('.')
                        $validOctets = $true
                        foreach ($octet in $octets) {
                            if ([int]$octet -lt 0 -or [int]$octet -gt 255) {
                                $validOctets = $false
                                break
                            }
                        }
                        
                        if (-not $validOctets) {
                            $specificIssue = "Invalid subnet format: Each octet must be between 0-255"
                            $subnetIssues += $specificIssue
                            $script:SubnetIssueDetails[$site.SiteCode] = $specificIssue
                            $shouldUpdateSubnet = $false
                        }

                    } else {
                        $specificIssue = "Invalid subnet format: Please use format like: XXX.XXX.XXX.XXX"
                        $subnetIssues += $specificIssue
                        $script:SubnetIssueDetails[$site.SiteCode] = $specificIssue
                        $shouldUpdateSubnet = $false
                    }
                    
                    # Check for duplicate subnets
                    if ($shouldUpdateSubnet) {
                        $existingEntries = $siteDataStore.GetAllEntries()
                        $existingSite = $existingEntries | Where-Object { $_.SiteCode -eq $site.SiteCode } | Select-Object -First 1
                        
                        # Check if subnet would create a duplicate
                        $duplicateSubnet = $existingEntries | Where-Object { 
                            $_.SiteSubnet -eq $site.SiteSubnet -and $_.SiteCode -ne $site.SiteCode 
                        } | Select-Object -First 1
                        
                        if ($duplicateSubnet) {
                            $specificIssue = "Subnet '$($site.SiteSubnet)' already exists for site '$($duplicateSubnet.SiteCode)'"
                            $subnetIssues += "$specificIssue - subnet update skipped"
                            $script:SubnetIssueDetails[$site.SiteCode] = $specificIssue
                            $shouldUpdateSubnet = $false
                        }
                    }
                }

                # [REST OF THE SITE PROCESSING CODE - SWITCHES, APs, etc. - SAME AS ORIGINAL]
                # SWITCHES DATA
                $switchCountText = Get-CellValue -Sheet $switchSheet -Row $row -Col 3
                if ([string]::IsNullOrWhiteSpace($switchCountText)) {
                    $site.SwitchCount = 1
                } else {
                    $site.SwitchCount = [Math]::Max(1, [Math]::Min(10, [int]$switchCountText))
                }

                $site.Switches.Clear()
                for ($i = 0; $i -lt $site.SwitchCount; $i++) {
                    $switch = [SwitchInfo]::new()
                    $colStart = 4 + ($i * 5)
                    
                    $switch.ManagementIP = Get-CellValue -Sheet $switchSheet -Row $row -Col $colStart
                    $switch.Name = Get-CellValue -Sheet $switchSheet -Row $row -Col ($colStart + 1)
                    $switch.AssetTag = Get-CellValue -Sheet $switchSheet -Row $row -Col ($colStart + 2)
                    $switch.Version = Get-CellValue -Sheet $switchSheet -Row $row -Col ($colStart + 3)
                    $switch.SerialNumber = Get-CellValue -Sheet $switchSheet -Row $row -Col ($colStart + 4)
                    
                    $site.Switches.Add($switch)
                }
                
                # ACCESS POINTS DATA
                $apCountText = Get-CellValue -Sheet $apSheet -Row $row -Col 3
                if ([string]::IsNullOrWhiteSpace($apCountText)) {
                    $site.APCount = 0
                } else {
                    $site.APCount = [Math]::Max(1, [Math]::Min(10, [int]$apCountText))
                }
                
                $site.AccessPoints.Clear()
                for ($i = 0; $i -lt $site.APCount; $i++) {
                    $ap = [AccessPointInfo]::new()
                    $colStart = 4 + ($i * 5)
                    
                    $ap.ManagementIP = Get-CellValue -Sheet $apSheet -Row $row -Col $colStart
                    $ap.Name = Get-CellValue -Sheet $apSheet -Row $row -Col ($colStart + 1)
                    $ap.AssetTag = Get-CellValue -Sheet $apSheet -Row $row -Col ($colStart + 2)
                    $ap.Version = Get-CellValue -Sheet $apSheet -Row $row -Col ($colStart + 3)
                    $ap.SerialNumber = Get-CellValue -Sheet $apSheet -Row $row -Col ($colStart + 4)
                    
                    $site.AccessPoints.Add($ap)
                }
                
                # FIREWALL DATA
                $site.FirewallIP = Get-CellValue -Sheet $firewallSheet -Row $row -Col 3
                $site.FirewallName = Get-CellValue -Sheet $firewallSheet -Row $row -Col 4
                $site.FirewallVersion = Get-CellValue -Sheet $firewallSheet -Row $row -Col 5
                $site.FirewallSN = Get-CellValue -Sheet $firewallSheet -Row $row -Col 6
                
                # PRIMARY CIRCUIT DATA
                $site.PrimaryCircuit.Vendor = Get-CellValue -Sheet $primarySheet -Row $row -Col 3
                $site.PrimaryCircuit.CircuitType = Get-CellValue -Sheet $primarySheet -Row $row -Col 4
                $site.PrimaryCircuit.CircuitID = Get-CellValue -Sheet $primarySheet -Row $row -Col 5
                $site.PrimaryCircuit.DownloadSpeed = Get-CellValue -Sheet $primarySheet -Row $row -Col 6
                $site.PrimaryCircuit.UploadSpeed = Get-CellValue -Sheet $primarySheet -Row $row -Col 7
                $site.PrimaryCircuit.IPAddress = Get-CellValue -Sheet $primarySheet -Row $row -Col 8
                $site.PrimaryCircuit.SubnetMask = Get-CellValue -Sheet $primarySheet -Row $row -Col 9
                $site.PrimaryCircuit.DefaultGateway = Get-CellValue -Sheet $primarySheet -Row $row -Col 10
                $site.PrimaryCircuit.DNS1 = Get-CellValue -Sheet $primarySheet -Row $row -Col 11
                $site.PrimaryCircuit.DNS2 = Get-CellValue -Sheet $primarySheet -Row $row -Col 12
                $site.PrimaryCircuit.RouterModel = Get-CellValue -Sheet $primarySheet -Row $row -Col 13
                $site.PrimaryCircuit.RouterName = Get-CellValue -Sheet $primarySheet -Row $row -Col 14
                $site.PrimaryCircuit.RouterSN = Get-CellValue -Sheet $primarySheet -Row $row -Col 15
                $site.PrimaryCircuit.PPPoEUsername = Get-CellValue -Sheet $primarySheet -Row $row -Col 16
                $site.PrimaryCircuit.PPPoEPassword = Get-CellValue -Sheet $primarySheet -Row $row -Col 17
                
                $hasModemText = Get-CellValue -Sheet $primarySheet -Row $row -Col 18
                $site.PrimaryCircuit.HasModem = ($hasModemText.ToLower() -eq "true")
                
                $site.PrimaryCircuit.ModemModel = Get-CellValue -Sheet $primarySheet -Row $row -Col 19
                $site.PrimaryCircuit.ModemName = Get-CellValue -Sheet $primarySheet -Row $row -Col 20
                $site.PrimaryCircuit.ModemSN = Get-CellValue -Sheet $primarySheet -Row $row -Col 21
                
                # BACKUP CIRCUIT DATA
                $hasBackupText = Get-CellValue -Sheet $backupSheet -Row $row -Col 3
                $site.HasBackupCircuit = ($hasBackupText.ToLower() -eq "true")
                
                if ($site.HasBackupCircuit) {
                    $site.BackupCircuit.Vendor = Get-CellValue -Sheet $backupSheet -Row $row -Col 4
                    $site.BackupCircuit.CircuitType = Get-CellValue -Sheet $backupSheet -Row $row -Col 5
                    $site.BackupCircuit.CircuitID = Get-CellValue -Sheet $backupSheet -Row $row -Col 6
                    $site.BackupCircuit.DownloadSpeed = Get-CellValue -Sheet $backupSheet -Row $row -Col 7
                    $site.BackupCircuit.UploadSpeed = Get-CellValue -Sheet $backupSheet -Row $row -Col 8
                    $site.BackupCircuit.IPAddress = Get-CellValue -Sheet $backupSheet -Row $row -Col 9
                    $site.BackupCircuit.SubnetMask = Get-CellValue -Sheet $backupSheet -Row $row -Col 10
                    $site.BackupCircuit.DefaultGateway = Get-CellValue -Sheet $backupSheet -Row $row -Col 11
                    $site.BackupCircuit.DNS1 = Get-CellValue -Sheet $backupSheet -Row $row -Col 12
                    $site.BackupCircuit.DNS2 = Get-CellValue -Sheet $backupSheet -Row $row -Col 13
                    $site.BackupCircuit.RouterModel = Get-CellValue -Sheet $backupSheet -Row $row -Col 14
                    $site.BackupCircuit.RouterName = Get-CellValue -Sheet $backupSheet -Row $row -Col 15
                    $site.BackupCircuit.RouterSN = Get-CellValue -Sheet $backupSheet -Row $row -Col 16
                    $site.BackupCircuit.PPPoEUsername = Get-CellValue -Sheet $backupSheet -Row $row -Col 17
                    $site.BackupCircuit.PPPoEPassword = Get-CellValue -Sheet $backupSheet -Row $row -Col 18
                    
                    $backupHasModemText = Get-CellValue -Sheet $backupSheet -Row $row -Col 19
                    $site.BackupCircuit.HasModem = ($backupHasModemText.ToLower() -eq "true")
                    
                    $site.BackupCircuit.ModemModel = Get-CellValue -Sheet $backupSheet -Row $row -Col 20
                    $site.BackupCircuit.ModemName = Get-CellValue -Sheet $backupSheet -Row $row -Col 21
                    $site.BackupCircuit.ModemSN = Get-CellValue -Sheet $backupSheet -Row $row -Col 22
                }
                
                # VLAN DATA
                $site.VLANs.VLAN100_Servers = Get-CellValue -Sheet $vlanSheet -Row $row -Col 3
                $site.VLANs.VLAN101_NetworkDevices = Get-CellValue -Sheet $vlanSheet -Row $row -Col 4
                $site.VLANs.VLAN102_UserDevices = Get-CellValue -Sheet $vlanSheet -Row $row -Col 5
                $site.VLANs.VLAN103_UserDevices2 = Get-CellValue -Sheet $vlanSheet -Row $row -Col 6
                $site.VLANs.VLAN104_VOIP = Get-CellValue -Sheet $vlanSheet -Row $row -Col 7
                $site.VLANs.VLAN105_WiFiCorp = Get-CellValue -Sheet $vlanSheet -Row $row -Col 8
                $site.VLANs.VLAN106_WiFiBYOD = Get-CellValue -Sheet $vlanSheet -Row $row -Col 9
                $site.VLANs.VLAN107_WiFiGuest = Get-CellValue -Sheet $vlanSheet -Row $row -Col 10
                $site.VLANs.VLAN108_Spare = Get-CellValue -Sheet $vlanSheet -Row $row -Col 11
                $site.VLANs.VLAN109_DMZ = Get-CellValue -Sheet $vlanSheet -Row $row -Col 12
                $site.VLANs.VLAN110_CCTV = Get-CellValue -Sheet $vlanSheet -Row $row -Col 13
                
                # CCTV DATA (15 cameras max)
                $cctvCountText = Get-CellValue -Sheet $cctvSheet -Row $row -Col 3
                if ([string]::IsNullOrWhiteSpace($cctvCountText)) {
                    $site.CCTVCount = 0
                } else {
                    $site.CCTVCount = [Math]::Max(0, [Math]::Min(15, [int]$cctvCountText))
                }
                
                $site.CCTVDevices.Clear()
                for ($i = 0; $i -lt $site.CCTVCount; $i++) {
                    $cctv = [CCTVInfo]::new()
                    $colStart = 4 + ($i * 3)
                    
                    $cctv.ManagementIP = Get-CellValue -Sheet $cctvSheet -Row $row -Col $colStart
                    $cctv.Name = Get-CellValue -Sheet $cctvSheet -Row $row -Col ($colStart + 1)
                    $cctv.SerialNumber = Get-CellValue -Sheet $cctvSheet -Row $row -Col ($colStart + 2)
                    
                    $site.CCTVDevices.Add($cctv)
                }
                
                # UPS DATA (5 UPS max)
                $upsCountText = Get-CellValue -Sheet $upsSheet -Row $row -Col 3
                if ([string]::IsNullOrWhiteSpace($upsCountText)) {
                    $site.UPSCount = 0
                } else {
                    $site.UPSCount = [Math]::Max(0, [Math]::Min(5, [int]$upsCountText))
                }
                
                $site.UPSDevices.Clear()
                for ($i = 0; $i -lt $site.UPSCount; $i++) {
                    $ups = [UPSInfo]::new()
                    $colStart = 4 + ($i * 2)
                    
                    $ups.ManagementIP = Get-CellValue -Sheet $upsSheet -Row $row -Col $colStart
                    $ups.Name = Get-CellValue -Sheet $upsSheet -Row $row -Col ($colStart + 1)
                    
                    $site.UPSDevices.Add($ups)
                }

                # PRINTER DATA (10 printers max)
                $printerCountText = Get-CellValue -Sheet $printerSheet -Row $row -Col 3
                if ([string]::IsNullOrWhiteSpace($printerCountText)) {
                    $site.PrinterCount = 0
                } else {
                    $site.PrinterCount = [Math]::Max(0, [Math]::Min(10, [int]$printerCountText))
                }

                $site.PrinterDevices.Clear()
                for ($i = 0; $i -lt $site.PrinterCount; $i++) {
                    $printer = [PrinterInfo]::new()
                    $colStart = 4 + ($i * 4)
                    
                    $printer.ManagementIP = Get-CellValue -Sheet $printerSheet -Row $row -Col $colStart
                    $printer.Name = Get-CellValue -Sheet $printerSheet -Row $row -Col ($colStart + 1)
                    $printer.Model = Get-CellValue -Sheet $printerSheet -Row $row -Col ($colStart + 2)
                    $printer.SerialNumber = Get-CellValue -Sheet $printerSheet -Row $row -Col ($colStart + 3)
                    
                    $site.PrinterDevices.Add($printer)
                }

                # VALIDATE IP ADDRESSES
                $invalidIPs = @()
                
                # Check firewall IP
                if (-not [string]::IsNullOrWhiteSpace($site.FirewallIP)) {
                    if (-not [ValidationUtility]::ValidateIP($site.FirewallIP)) {
                        $invalidIPs += "Firewall IP: $($site.FirewallIP)"
                    }
                }
                
                # Check circuit IPs
                if (-not [string]::IsNullOrWhiteSpace($site.PrimaryCircuit.IPAddress)) {
                    if (-not [ValidationUtility]::ValidateIP($site.PrimaryCircuit.IPAddress)) {
                        $invalidIPs += "Primary Circuit IP: $($site.PrimaryCircuit.IPAddress)"
                    }
                }
                
                if (-not [string]::IsNullOrWhiteSpace($site.BackupCircuit.IPAddress)) {
                    if (-not [ValidationUtility]::ValidateIP($site.BackupCircuit.IPAddress)) {
                        $invalidIPs += "Backup Circuit IP: $($site.BackupCircuit.IPAddress)"
                    }
                }
                
                # Check device IPs
                foreach ($switch in $site.Switches) {
                    if (-not [string]::IsNullOrWhiteSpace($switch.ManagementIP)) {
                        if (-not [ValidationUtility]::ValidateIP($switch.ManagementIP)) {
                            $invalidIPs += "Switch IP: $($switch.ManagementIP)"
                        }
                    }
                }
                
                foreach ($ap in $site.AccessPoints) {
                    if (-not [string]::IsNullOrWhiteSpace($ap.ManagementIP)) {
                        if (-not [ValidationUtility]::ValidateIP($ap.ManagementIP)) {
                            $invalidIPs += "Access Point IP: $($ap.ManagementIP)"
                        }
                    }
                }
                
                foreach ($cctv in $site.CCTVDevices) {
                    if (-not [string]::IsNullOrWhiteSpace($cctv.ManagementIP)) {
                        if (-not [ValidationUtility]::ValidateIP($cctv.ManagementIP)) {
                            $invalidIPs += "CCTV IP: $($cctv.ManagementIP)"
                        }
                    }
                }
                
                foreach ($ups in $site.UPSDevices) {
                    if (-not [string]::IsNullOrWhiteSpace($ups.ManagementIP)) {
                        if (-not [ValidationUtility]::ValidateIP($ups.ManagementIP)) {
                            $invalidIPs += "UPS IP: $($ups.ManagementIP)"
                        }
                    }
                }

                foreach ($printer in $site.PrinterDevices) {
                    if (-not [string]::IsNullOrWhiteSpace($printer.ManagementIP)) {
                        if (-not [ValidationUtility]::ValidateIP($printer.ManagementIP)) {
                            $invalidIPs += "Printer IP: $($printer.ManagementIP)"
                        }
                    }
                }

                # If there are invalid IPs, throw error (this still blocks everything)
                if ($invalidIPs.Count -gt 0) {
                    throw "Invalid IP addresses: $($invalidIPs -join ', ')"
                }

                # NOW HANDLE DUPLICATES AND SAVE - WITH SELECTIVE SUBNET UPDATE
                $existingEntries = $siteDataStore.GetAllEntries()
                $existingSite = $existingEntries | Where-Object { $_.SiteCode -eq $site.SiteCode } | Select-Object -First 1

                if ($existingSite) {
                    # EXISTING SITE - Update everything except problematic subnet
                    if (-not $shouldUpdateSubnet -and $subnetIssues.Count -gt 0) {                            
                        # Keep the original subnet
                        $site.SiteSubnet = $existingSite.SiteSubnet
                    }
                    
                    # Proceed with update logic
                    switch ($importMode) {
                        "Skip" {
                            $skippedCount++
                            $script:ProcessedSites += @{SiteCode = $site.SiteCode; Status = "Skipped"}
                        }
                        "Update" {
                            # Smart update - only update fields with new data
                            $updateResult = Update-SiteWithNewData -ExistingSite $existingSite -ImportSite $site
                            
                            if ($updateResult.HasChanges) {
                                if ($siteDataStore.UpdateEntry($updateResult.Site)) {
                                    $importedCount++
                                    
                                    # FIXED: Properly format subnet warning
                                    $statusDetails = [System.Collections.ArrayList]::new()
                                    
                                    # Add regular change details
                                    foreach ($detail in $updateResult.ChangeDetails) {
                                        $statusDetails.Add($detail) | Out-Null
                                    }
                                    
                                    # Add subnet warning separately if applicable
                                    if ($subnetIssues.Count -gt 0) {
                                        $cleanWarning = $subnetIssues[0] -replace " - subnet update skipped", ""
                                        $statusDetails.Add("WARNING: Subnet not updated - $cleanWarning") | Out-Null
                                    }
                                    
                                    $script:ProcessedSites += @{
                                        SiteCode = $site.SiteCode; 
                                        Status = "Updated"; 
                                        Changes = $updateResult.ChangesCount; 
                                        Details = $statusDetails.ToArray()
                                    }
                                }
                            } else {
                                # Check if only issue was subnet
                                if ($subnetIssues.Count -gt 0) {
                                    $skippedCount++
                                    $script:ProcessedSites += @{
                                        SiteCode = $site.SiteCode; 
                                        Status = "SubnetIssueOnly"; 
                                        Details = @("Subnet validation prevented updates: $($subnetIssues[0])")
                                    }
                                } else {
                                    $skippedCount++
                                    $script:ProcessedSites += @{SiteCode = $site.SiteCode; Status = "NoChanges"}
                                }
                            }
                        }
                        "Replace" {
                            # Complete replacement but keep original subnet if problematic
                            if (-not $shouldUpdateSubnet) {
                                $site.SiteSubnet = $existingSite.SiteSubnet
                            }
                            
                            $site.ID = $existingSite.ID  # Keep the same ID
                            if ($siteDataStore.UpdateEntry($site)) {
                                $importedCount++
                                $script:ProcessedSites += @{SiteCode = $site.SiteCode; Status = "Updated"}
                            }
                        }
                    }

                } else {
                    # NEW SITE - Subnet issues block new site creation
                    if ($subnetIssues.Count -gt 0) {
                        throw $subnetIssues -join '; '
                    }
                    
                    # New site - add normally
                    if ($siteDataStore.AddEntry($site)) {
                        $importedCount++
                        $script:ProcessedSites += @{SiteCode = $site.SiteCode; Status = "New"}
                    }
                }
                
                # Update progress
                $progress = [math]::Round((($row - 1) / $totalSites) * 90) + 10
                $pbSiteImportProgress.Value = $progress
                $txtSiteProgressDetails.Text = "$importedCount / $totalSites sites processed (Current: $($site.SiteCode))"
                [System.Windows.Forms.Application]::DoEvents()
                
            } catch {
               $skippedCount++
               $siteCodeForError = try { Get-CellValue -Sheet $basicSheet -Row $row -Col 2 } catch { "Row$row" }
               $errorMessages += "Row $row (Site: $siteCodeForError): $_"
               $script:ProcessedSites += @{SiteCode = $siteCodeForError; Status = "Error"}
               continue
           }
       }
       
       # Build comprehensive result message
       $result = @()
       $result += "Excel import completed successfully!"
       $result += "=========================================="
       $result += "Total sites processed: $totalSites"
       $result += ""

       # Count different outcomes INCLUDING subnet warnings
       $newSitesCount = 0
       $updatedSitesCount = 0 
       $skippedDuplicatesCount = 0
       $errorSitesCount = 0
       $noChangesCount = 0
       $subnetWarningsCount = $script:SubnetIssueDetails.Count

       # Arrays to track sites by outcome - FIXED
       $newSites = @()
       $updatedSites = @()
       $skippedSites = @()
       $errorSites = @()
       $noChangesSites = @()

       # Process the tracking arrays - PREVENT DOUBLE COUNTING
       foreach ($site in $script:ProcessedSites) {
           switch ($site.Status) {
               "New" { 
                   $newSites += $site.SiteCode 
                   $newSitesCount++
               }
               "Updated" { 
                   $updatedSites += $site.SiteCode 
                   $updatedSitesCount++
                   # DO NOT add to noChangesSites even if it has subnet warnings
               }
               "Skipped" { 
                   $skippedSites += $site.SiteCode 
                   $skippedDuplicatesCount++
               }
               "NoChanges" { 
                   # ONLY add to noChangesSites if NOT already in updatedSites
                   if ($site.SiteCode -notin $updatedSites) {
                       $noChangesSites += $site.SiteCode 
                       $noChangesCount++
                   }
               }
               "SubnetIssueOnly" {
                   # Sites with ONLY subnet issues (no other changes) - count as no changes
                   if ($site.SiteCode -notin $updatedSites) {
                       $noChangesSites += $site.SiteCode
                       $noChangesCount++
                   }
               }
               "Error" { 
                   $errorSites += $site.SiteCode 
                   $errorSitesCount++
               }
           }
       }

       # Remove duplicates from noChangesSites (final safety check)
       $noChangesSites = $noChangesSites | Sort-Object | Get-Unique
       $noChangesCount = $noChangesSites.Count

       # Summary counts - FIXED
       if ($newSites.Count -gt 0) {
           $result += " Successfully imported: $($newSites.Count) sites"
       }
       if ($updatedSites.Count -gt 0) {
           $result += " Updated existing: $($updatedSites.Count) sites"
       }
       if ($noChangesSites.Count -gt 0) {
           $result += " No changes needed: $($noChangesSites.Count) sites"
       }
       if ($skippedSites.Count -gt 0) {
           $result += " Skipped duplicates: $($skippedSites.Count) sites"
       }
       if ($subnetWarningsCount -gt 0) {
           $result += " Subnet warnings: $subnetWarningsCount sites"
       }
       if ($errorSites.Count -gt 0) {
           $result += " Failed validation: $($errorSites.Count) sites"
       }

       $result += ""

       # Detailed site listings
       if ($updatedSites.Count -gt 0) {
           $result += "UPDATED:"
           foreach ($site in $script:ProcessedSites) {
               if ($site.Status -eq "Updated") {
                   # FILTER OUT subnet warnings from the details
                   $filteredDetails = @()
                   foreach ($detail in $site.Details) {
                       if ($detail -notlike "*WARNING: Subnet not updated*") {
                           $filteredDetails += $detail
                       }
                   }
                   
                   if ($filteredDetails.Count -gt 0) {
                       $changesList = $filteredDetails -join ", "
                       $result += "$($site.SiteCode): $changesList"
                   } else {
                       # If only subnet warnings were removed, show that fields were updated
                       $result += "$($site.SiteCode): Updated (see subnet warnings below)"
                   }
               }
           }
           $result += ""
       }

       if ($noChangesSites.Count -gt 0) {
           $result += "NO CHANGES NEEDED:"
           # Join all sites with commas on one line
           $uniqueNoChanges = $noChangesSites | Sort-Object | Get-Unique
           $result += ($uniqueNoChanges -join ", ")
           $result += ""
       }

       if ($script:SubnetIssueDetails.Count -gt 0) {
           $result += "SUBNET WARNINGS:"
           
           # Sort the sites alphabetically and show their specific issues
           $sortedSites = $script:SubnetIssueDetails.Keys | Sort-Object
           
           foreach ($siteCode in $sortedSites) {
               $specificMessage = $script:SubnetIssueDetails[$siteCode]
               $result += "$siteCode`: $specificMessage"
           }
           
           $result += ""
       }

       if ($errorSites.Count -gt 0) {
           $result += "VALIDATION ERRORS:"
           foreach ($errorSite in $errorSites) {
               $errorDetail = $errorMessages | Where-Object { $_ -like "*$errorSite*" } | Select-Object -First 1
               if ($errorDetail) {
                   $result += "❌ $errorSite`: $($errorDetail -replace '^.*?: ', '')"
               } else {
                   $result += "❌ $errorSite`: Validation failed"
               }
           }
           $result += ""
       }

       if ($importedCount -gt 0) {
           $result += "Data grid refreshed with changes applied."
       }

       return ($result -join "`n")
       
   } catch {
       throw "Excel import failed: $_"
   }
   
    finally {
        # CRITICAL: Enhanced cleanup to eliminate all Excel processes
        
        # Release all worksheet references first
        if ($basicSheet) { $basicSheet = Release-ComObject $basicSheet }
        if ($switchSheet) { $switchSheet = Release-ComObject $switchSheet }
        if ($apSheet) { $apSheet = Release-ComObject $apSheet }
        if ($firewallSheet) { $firewallSheet = Release-ComObject $firewallSheet }
        if ($primarySheet) { $primarySheet = Release-ComObject $primarySheet }
        if ($backupSheet) { $backupSheet = Release-ComObject $backupSheet }
        if ($vlanSheet) { $vlanSheet = Release-ComObject $vlanSheet }
        if ($cctvSheet) { $cctvSheet = Release-ComObject $cctvSheet }
        if ($upsSheet) { $upsSheet = Release-ComObject $upsSheet }
        
        # Close and release workbook
        if ($workbook) {
            try { 
                $workbook.Close($false)  # Close without saving
            } catch { }
            $workbook = Release-ComObject $workbook
        }
        
        # Quit and release Excel application
        if ($excel) {
            try {
                $excel.DisplayAlerts = $true
                $excel.Quit()
            } catch { }
            $excel = Release-ComObject $excel
        }
        
        # Multiple rounds of garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Small delay to ensure process termination
        Start-Sleep -Milliseconds 500
        
        # Hide progress panel
        $pnlSiteImportProgress.Visibility = [System.Windows.Visibility]::Collapsed
    }
}