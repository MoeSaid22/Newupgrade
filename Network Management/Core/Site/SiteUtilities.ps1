# SiteUtilities.ps1 - Utility functions for Site Management
# Note: This file must be loaded AFTER WPF assemblies and BEFORE SiteModels.ps1

# ===================================================================
# UTILITY FUNCTIONS AND CLASSES
# ===================================================================

# Helper function for showing messages that works in both WPF and non-WPF environments
function Show-MessageBox {
    param([string]$Message, [string]$Title = "Message", [string]$Button = "OK", [string]$Icon = "Information")
    
    try {
        Show-MessageBox $Message $Title $Button $Icon
    } catch {
        # Fallback for non-WPF environments
        Write-Host "$Title`: $Message" -ForegroundColor $(
            switch ($Icon) {
                "Error" { "Red" }
                "Warning" { "Yellow" }
                default { "White" }
            }
        )
    }
}

# Safely release COM objects to prevent memory leaks
function Safe-ReleaseComObject {
    param([object]$ComObject)
    if ($ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null } 
        catch { Write-Warning "Failed to release COM object: $_" }
    }
}

# Format phone number to standard format
function Format-PhoneNumber {
    param([string]$PhoneNumber)
    
    if ([string]::IsNullOrWhiteSpace($PhoneNumber)) { return "" }
    
    # Remove all non-digits
    $digits = $PhoneNumber -replace '[^\d]', ''
    
    # Skip formatting if not exactly 10 digits
    if ($digits.Length -ne 10) { return $PhoneNumber }
    
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

# Basic validation utility class for IP addresses - SiteEntry validation moved to SiteModels.ps1
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
}