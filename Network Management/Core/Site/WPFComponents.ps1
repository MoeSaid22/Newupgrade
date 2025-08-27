# WPFComponents.ps1 - WPF-specific components (requires WPF assemblies)

# Phone number converter for XAML binding
class PhoneNumberConverter : System.Windows.Data.IValueConverter {
    [object] Convert([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return Format-PhoneNumber $value
    }
    
    [object] ConvertBack([object]$value, [System.Type]$targetType, [object]$parameter, [System.Globalization.CultureInfo]$culture) {
        return $value
    }
}