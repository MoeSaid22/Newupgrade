# Main.ps1 - Network Management Tool Entry Point

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
$mainScript = Join-Path $scriptPath $appFolder | Join-Path -ChildPath "Core" | Join-Path -ChildPath "Site" | Join-Path -ChildPath "Site.ps1"

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

# Load required class definitions - Now using Site.ps1 that contains all models
# Load WPF assemblies first (needed for PhoneNumberConverter)
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue
    Add-Type -AssemblyName WindowsBase -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
} catch {
    Write-Host "WARNING: Could not load WPF assemblies (expected in non-Windows environments)" -ForegroundColor Yellow
}

$coreScript = Join-Path $appFolderPath "Core" | Join-Path -ChildPath "Site" | Join-Path -ChildPath "Site.ps1"
if (-not (Test-Path -Path $coreScript -PathType Leaf)) {
    Write-Host "ERROR: Core Site script not found at $coreScript" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    . $coreScript
    Write-Host "Successfully loaded Site script with class definitions" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to load Site script: $_" -ForegroundColor Red
    Write-Host "Please check Site.ps1 for syntax errors (missing braces, etc.)" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Import all required modules from their new locations
$coreModules = @{
    "SiteImportExport.ps1" = "Site"
    "IPNetworkModule.ps1" = "IP" 
    "DeviceManager.ps1" = "Site"
    "EditSiteWindow.ps1" = "Site"
}

foreach ($moduleInfo in $coreModules.GetEnumerator()) {
    $module = $moduleInfo.Key
    $subFolder = $moduleInfo.Value
    $modulePath = Join-Path $appFolderPath "Core" | Join-Path -ChildPath $subFolder | Join-Path -ChildPath $module
    if (-not (Test-Path $modulePath -PathType Leaf)) {
        Write-Host "WARNING: Module not found at $modulePath" -ForegroundColor Yellow
        continue
    }
    
    try {
        . $modulePath
        Write-Host "Loaded module: $module from Core/$subFolder/" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to load module $module : $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Check XAML file
$xamlPath = Join-Path $appFolderPath "UI" | Join-Path -ChildPath "NetworkManagement.xaml"
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
    # Note: Site.ps1 is already loaded above, don't load it again
    Write-Host "Application loaded successfully!" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to run main script: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    Read-Host "Press Enter to exit"
    exit 1
}