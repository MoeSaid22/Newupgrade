# Load required WPF assemblies FIRST - before anything else!
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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
        Write-Host "Warning: Could not create shortcut ($_)" -ForegroundColor Yellow
    }
}

# Set execution policy for this session only
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue | Out-Null

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

# Load required class definitions
$importExportScript = Join-Path $appFolderPath "Core" | Join-Path -ChildPath "DataModels.ps1"
if (-not (Test-Path -Path $importExportScript -PathType Leaf)) {
    Write-Host "ERROR: DataModels script not found at $importExportScript" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    . $importExportScript
    Write-Host "Successfully loaded class definitions" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to load class definitions: $_" -ForegroundColor Red
    Write-Host "Please check DataModels.ps1 for syntax errors (missing braces, etc.)" -ForegroundColor Yellow
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
        Write-Host "WARNING: Module not found at $modulePath" -ForegroundColor Yellow
        continue
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
$xamlPath = Join-Path (Split-Path $mainScript -Parent) ".." | Join-Path -ChildPath "UI" | Join-Path -ChildPath "NetworkManagement.xaml"
if (-not (Test-Path $xamlPath -PathType Leaf)) {
    Write-Host "ERROR: XAML file not found at $xamlPath" -ForegroundColor Red
    Write-Host "Please check that the file exists and the path is correct" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    [xml]$null = Get-Content $xamlPath -Raw
    Write-Host "XAML file syntax is valid" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Invalid XAML syntax: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Run the main script
try {
    Write-Host "Starting main application..." -ForegroundColor Green
    . $mainScript
}
catch {
    Write-Host "ERROR: Failed to run main script: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    Read-Host "Press Enter to exit"
    exit 1
}