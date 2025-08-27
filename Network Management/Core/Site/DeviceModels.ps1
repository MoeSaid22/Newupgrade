# DeviceModels.ps1 - Device class definitions for Network Management

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