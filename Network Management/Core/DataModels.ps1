# DataModels.ps1 - Data model class definitions for Network Management

# Switch data structure
class SwitchInfo {
    [string]$ManagementIP
    [string]$Name
    [string]$AssetTag
    [string]$Version
    [string]$SerialNumber
}

# Access Point data structure
class AccessPointInfo {
    [string]$ManagementIP
    [string]$Name
    [string]$AssetTag
    [string]$Version
    [string]$SerialNumber
}

# UPS data structure
class UPSInfo {
    [string]$ManagementIP
    [string]$Name
    [string]$AssetTag
    [string]$Version
    [string]$SerialNumber
}

# CCTV data structure
class CCTVInfo {
    [string]$ManagementIP
    [string]$Name
    [string]$SerialNumber
}

class PrinterInfo {
    [string]$ManagementIP
    [string]$Name
    [string]$Model
    [string]$SerialNumber
}

# Circuit data structure
class CircuitInfo {
    [string]$Vendor
    [string]$CircuitType
    [string]$CircuitID
    [string]$DownloadSpeed
    [string]$UploadSpeed
    [string]$IPAddress
    [string]$SubnetMask
    [string]$DefaultGateway
    [string]$DNS1
    [string]$DNS2
    [string]$RouterModel
    [string]$RouterName
    [string]$RouterSN
    [string]$PPPoEUsername
    [string]$PPPoEPassword
    [bool]$HasModem
    [string]$ModemModel
    [string]$ModemName
    [string]$ModemSN
}

# VLAN data structure
class VLANInfo {
    [string]$VLAN100_Servers
    [string]$VLAN101_NetworkDevices
    [string]$VLAN102_UserDevices
    [string]$VLAN103_UserDevices2
    [string]$VLAN104_VOIP
    [string]$VLAN105_WiFiCorp
    [string]$VLAN106_WiFiBYOD
    [string]$VLAN107_WiFiGuest
    [string]$VLAN108_Spare
    [string]$VLAN109_DMZ
    [string]$VLAN110_CCTV
}

# Main Site Entry Class
class SiteEntry {
    [int]$ID
    # Basic Info
    [string]$SiteCode
    [string]$SiteSubnet
    [string]$SiteSubnetCode
    [string]$SiteName
    [string]$SiteAddress
    [string]$MainContactName
    [string]$MainContactPhone
    [string]$SecondContactName
    [string]$SecondContactPhone


    [string]$MainContactPhoneFormatted
    [string]$SecondContactPhoneFormatted
    
    # Network Equipment
    [int]$SwitchCount
    [System.Collections.Generic.List[SwitchInfo]]$Switches
    [int]$APCount
    [System.Collections.Generic.List[AccessPointInfo]]$AccessPoints
    [int]$UPSCount
    [System.Collections.Generic.List[UPSInfo]]$UPSDevices
    [int]$CCTVCount
    [System.Collections.Generic.List[CCTVInfo]]$CCTVDevices
    [int]$PrinterCount
    [System.Collections.Generic.List[PrinterInfo]]$PrinterDevices
    [string]$FirewallIP
    [string]$FirewallName
    [string]$FirewallVersion
    [string]$FirewallSN
    
    # Circuits
    [CircuitInfo]$PrimaryCircuit
    [bool]$HasBackupCircuit
    [CircuitInfo]$BackupCircuit
    
    # VLANs
    [VLANInfo]$VLANs

    # Properties for DataGrid display
    [string]$Switch1IP
    [string]$Switch1Name
    [string]$PrimaryVendor
    [string]$PrimaryCircuitIP
    [string]$PrimaryDownloadSpeed
    [string]$PrimaryUploadSpeed
    [string]$BackupVendor
    [string]$BackupCircuitIP
    [string]$BackupDownloadSpeed
    [string]$BackupUploadSpeed

    SiteEntry() {
        $this.SwitchCount = 1
        $this.Switches = [System.Collections.Generic.List[SwitchInfo]]::new()
        $this.APCount = 1
        $this.AccessPoints = [System.Collections.Generic.List[AccessPointInfo]]::new()
        $this.UPSCount = 1
        $this.UPSDevices = [System.Collections.Generic.List[UPSInfo]]::new()
        $this.CCTVCount = 1
        $this.CCTVDevices = [System.Collections.Generic.List[CCTVInfo]]::new()
        $this.PrinterCount = 1
        $this.PrinterDevices = [System.Collections.Generic.List[PrinterInfo]]::new()
        $this.PrimaryCircuit = [CircuitInfo]::new()
        $this.HasBackupCircuit = $false
        $this.BackupCircuit = [CircuitInfo]::new()
        $this.VLANs = [VLANInfo]::new()
    }

    [void] UpdateDisplayProperties() {
    $this.Switch1IP = if ($this.Switches.Count -gt 0) { $this.Switches[0].ManagementIP } else { "" }
    $this.Switch1Name = if ($this.Switches.Count -gt 0) { $this.Switches[0].Name } else { "" }
    $this.PrimaryVendor = $this.PrimaryCircuit.Vendor
    $this.PrimaryCircuitIP = $this.PrimaryCircuit.IPAddress
    $this.PrimaryDownloadSpeed = $this.PrimaryCircuit.DownloadSpeed
    $this.PrimaryUploadSpeed = $this.PrimaryCircuit.UploadSpeed
    $this.BackupVendor = $this.BackupCircuit.Vendor
    $this.BackupCircuitIP = $this.BackupCircuit.IPAddress
    $this.BackupDownloadSpeed = $this.BackupCircuit.DownloadSpeed
    $this.BackupUploadSpeed = $this.BackupCircuit.UploadSpeed
    
    # Format phone numbers for display WITHOUT modifying original data
    $this.MainContactPhoneFormatted = Format-PhoneNumber $this.MainContactPhone
    $this.SecondContactPhoneFormatted = Format-PhoneNumber $this.SecondContactPhone
    }
}

# ===================================================================
# UTILITY FUNCTIONS AND CLASSES
# ===================================================================

# Safely release COM objects to prevent memory leaks
function Release-ComObject {
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [System.Object]$ComObject
    )
    
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

# Get safe string value from object, returns empty string if null
function Get-SafeValue {
    param([object]$Value)
    if ($Value) { return $Value.ToString() } else { return "" }
}

# Phone number converter for XAML binding
class PhoneNumberConverter : System.Windows.Data.IValueConverter {
    [object] Convert([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return Format-PhoneNumber $value
    }
    
    [object] ConvertBack([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return $value
    }
}

# Validation utility class for IP addresses and data validation
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
# Site data store class for managing persistent site data
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
