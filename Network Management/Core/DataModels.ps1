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