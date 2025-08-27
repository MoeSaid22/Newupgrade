# Check if shortcut exists, create if missing
$shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "Network Management Tool.lnk")
if (-not (Test-Path $shortcutPath)) {
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        
        # Point to PowerShell executable (hidden window)
        $shortcut.TargetPath = "powershell.exe"
        
        # Arguments to run your script silently
        $shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
        
        # Set icon (optional - use your own .ico file or PowerShell's)
        $shortcut.IconLocation = "powershell.exe,0"
        
        $shortcut.WorkingDirectory = Split-Path $MyInvocation.MyCommand.Path
        $shortcut.Description = "Network Management Tool"
        $shortcut.Save()
        
        Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not create shortcut ($($_.Exception.Message))" -ForegroundColor Yellow
    }
}

# Main.ps1 - Launcher script with execution policy bypass

# Set execution policy for this session only (Windows only)
try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop | Out-Null
}
catch {
    # Ignore execution policy errors on non-Windows platforms
    Write-Host "Note: Execution policy setting not available on this platform" -ForegroundColor Yellow
}

# Get the directory where this script is located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path to the main script
$appFolder = "Network Management"
$mainScript = Join-Path $scriptPath $appFolder | Join-Path -ChildPath "Core" | Join-Path -ChildPath "Site.ps1"

# Check if the app folder and main script exist
$appFolderPath = Join-Path $scriptPath $appFolder
if (-not (Test-Path -Path $appFolderPath -PathType Container)) {
    Write-Host "ERROR: App folder not found at $appFolderPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not (Test-Path -Path $mainScript -PathType Leaf)) {
    Write-Host "ERROR: Main script not found at $mainScript" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Load required assemblies first (before any class definitions that depend on them)
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop | Out-Null
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop | Out-Null
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop | Out-Null
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
    Write-Host "Loaded required assemblies" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to load required assemblies: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Import all required modules
$coreModules = @(
    "DataModels.ps1",
    "SiteImportExport.ps1", 
    "IPNetworkModule.ps1",
    "DeviceManager.ps1",
    "EditSiteWindow.ps1"
)

foreach ($module in $coreModules) {
    $modulePath = Join-Path $appFolderPath "Core" | Join-Path -ChildPath $module
    if (-not (Test-Path $modulePath -PathType Leaf)) {
        Write-Host "ERROR: Module not found at $modulePath" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    try {
        . $modulePath
        Write-Host "Loaded module: $module" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to load module $module : $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Check XAML file
$xamlPath = Join-Path (Split-Path $mainScript -Parent) ".." | Join-Path -ChildPath "UI" | Join-Path -ChildPath "SiteNetworkIdentifier.xaml"
if (-not (Test-Path $xamlPath -PathType Leaf)) {
    Write-Host "ERROR: XAML file not found at $xamlPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    [xml]$null = Get-Content $xamlPath -Raw
}
catch {
    Write-Host "ERROR: Invalid XAML syntax: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Run the main script
try {
    . $mainScript
}
catch {
    Write-Host "ERROR: Failed to run main script: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    Read-Host "Press Enter to exit"
    exit 1
}