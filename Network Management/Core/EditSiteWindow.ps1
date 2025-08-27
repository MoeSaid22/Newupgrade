# EditSiteWindow.ps1 - Edit site window functions for Network Management

# Import required modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import data models
$dataModelsPath = Join-Path $scriptPath "DataModels.ps1"
if (Test-Path $dataModelsPath) {
    . $dataModelsPath
}

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
