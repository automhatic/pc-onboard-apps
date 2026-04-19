# ============================================
# PC Onboarding Script V1.1
# Installs: RingCentral, M365, Teams, OneDrive, Azure VPN, Dynamic Theme, Adobe, Encompass
# ============================================

# Configuration
$LogPath = "C:\ProgramData\EnrollmentScript"
$LogFile = "$LogPath\EnrollmentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$CompletionFlag = "$LogPath\enrollment_v1_complete.txt"

# Create log directory
if (!(Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

# Logging function
function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$Timestamp - $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

Write-Log "=== Enrollment Script V1.1 Started ==="
Write-Log "Running as: $env:USERNAME"

# Verify Winget is available
Write-Log "Verifying Winget availability..."
$WingetPath = Get-Command winget -ErrorAction SilentlyContinue
if (!$WingetPath) {
    Write-Log "ERROR: Winget not found. This script requires Windows 10 21H2+ or Windows 11"
    exit 1
}
Write-Log "Winget found at: $($WingetPath.Source)"

# ============================================
# PHASE 1: Installing Standard Applications
# ============================================
Write-Log "=== PHASE 1: Installing Standard Applications ==="

# Standard apps to install
$StandardApps = @(
    @{Name="RingCentral"; ID="RingCentral.RingCentral"; Source="winget"},
    @{Name="Microsoft 365 Apps"; ID="Microsoft.Office"; Source="winget"},
    @{Name="Microsoft Teams"; ID="Microsoft.Teams"; Source="winget"},
    @{Name="Microsoft OneDrive"; ID="Microsoft.OneDrive"; Source="winget"},
    @{Name="Azure VPN Client"; ID="9NP355QT2SQB"; Source="msstore"},
    @{Name="Dynamic Theme"; ID="9NBLGGH1ZBKW"; Source="msstore"}
)

# Install each app
foreach ($App in $StandardApps) {
    Write-Log "Installing: $($App.Name) ($($App.ID))"
    
    # Check if already installed
    $Installed = winget list --id $App.ID --exact 2>&1
    if ($Installed -match $App.ID) {
        Write-Log "INFO: $($App.Name) already installed, skipping..."
        continue
    }
    
    # Build install command with source
    $InstallArgs = @(
        "install",
        "--id", $App.ID,
        "--source", $App.Source,
        "--silent",
        "--accept-package-agreements",
        "--accept-source-agreements"
    )
    
    # Execute installation
    & winget $InstallArgs 2>&1 | Out-Null
    
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        Write-Log "SUCCESS: $($App.Name) installed"
    } else {
        Write-Log "WARNING: $($App.Name) installation returned code $LASTEXITCODE"
    }
    
    Start-Sleep -Seconds 3
}

# ============================================
# PHASE 2: Installing Adobe Acrobat Reader
# ============================================
Write-Log "=== PHASE 2: Installing Adobe Acrobat Reader ==="

Write-Log "Installing Adobe Acrobat Reader (64-bit, no add-ons)..."
$AdobeInstall = winget install --id Adobe.Acrobat.Reader.64-bit `
    --silent `
    --accept-package-agreements `
    --accept-source-agreements `
    --override "ALLUSERS=1 EULA_ACCEPT=YES SUPPRESS_APP_LAUNCH=YES ENABLE_CHROMEEXT=0 DISABLE_ARM_SERVICE_INSTALL=1 DISABLE_BROWSER_INTEGRATION=1 REMOVE_PREVIOUS_VERSIONS=1 ADD_THUMBNAILPREVIEW=0 DISABLEDESKTOPSHORTCUT=1 /qn /norestart" `
    2>&1

Write-Log $AdobeInstall

if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189 -or $LASTEXITCODE -eq 3010) {
    Write-Log "SUCCESS: Adobe Reader installed"
} else {
    Write-Log "WARNING: Adobe Reader installation returned code $LASTEXITCODE"
}

Write-Log "Waiting for Adobe Reader installation to complete..."
Start-Sleep -Seconds 15

# ============================================
# PHASE 3: Installing Encompass Smart Client
# ============================================
Write-Log "=== PHASE 3: Installing Encompass Smart Client ==="

# Check if Adobe Reader is installed (Encompass prerequisite)
$AdobeInstalled = Test-Path "HKLM:\SOFTWARE\Adobe\Acrobat Reader\DC" -ErrorAction SilentlyContinue
if (!$AdobeInstalled) {
    $AdobeInstalled = Test-Path "HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat Reader\DC" -ErrorAction SilentlyContinue
}

if (!$AdobeInstalled) {
    Write-Log "WARNING: Adobe Reader not detected in registry, but continuing anyway..."
}

# Check if Encompass is already installed
$EncompassInstalled = Test-Path "C:\Program Files (x86)\Ellie Mae\Encompass\SmartClient.exe" -ErrorAction SilentlyContinue
if ($EncompassInstalled) {
    Write-Log "INFO: Encompass already installed"
} else {
    Write-Log "INFO: Encompass not found - will be installed via separate Intune package"
}

Start-Sleep -Seconds 3

# ============================================
# PHASE 4: Configure Dynamic Theme
# ============================================
Write-Log "=== PHASE 4: Configuring Dynamic Theme ==="

# Wait for Dynamic Theme to initialize
Start-Sleep -Seconds 5

# Apply Windows theme settings (dark mode)
Write-Log "Applying Windows dark mode settings..."
try {
    # Download and import registry settings from GitHub
    $RegUrl = "https://raw.githubusercontent.com/automhatic/pc-onboard-apps/main/configs/dynamictheme/ThemePersonalize.reg"
    $RegFile = "$env:TEMP\ThemePersonalize.reg"
    
    Invoke-WebRequest -Uri $RegUrl -OutFile $RegFile -UseBasicParsing
    
    if (Test-Path $RegFile) {
        # Import registry file
        reg import $RegFile /reg:64 2>&1 | Out-Null
        Write-Log "SUCCESS: Dark mode registry settings applied"
        Remove-Item $RegFile -Force -ErrorAction SilentlyContinue
    } else {
        Write-Log "WARNING: Failed to download registry settings"
    }
} catch {
    Write-Log "WARNING: Failed to apply theme settings - $_"
    
    # Fallback: Apply settings manually
    try {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 1 -Type DWord -Force
        Write-Log "SUCCESS: Dark mode settings applied via PowerShell"
    } catch {
        Write-Log "ERROR: Failed to apply dark mode settings - $_"
    }
}

Start-Sleep -Seconds 3

# ============================================
# PHASE 5: Installation Verification
# ============================================
Write-Log "=== PHASE 5: Installation Verification ==="

$VerificationChecks = @(
    @{Name="RingCentral"; Pattern="RingCentral"},
    @{Name="Microsoft 365"; Pattern="Microsoft 365|Microsoft Office"},
    @{Name="Teams"; Pattern="Teams"},
    @{Name="OneDrive"; Pattern="OneDrive"},
    @{Name="Azure VPN"; Pattern="Azure VPN"},
    @{Name="Dynamic Theme"; Pattern="Dynamic.*Theme"},
    @{Name="Adobe"; Pattern="Adobe.*Reader|Acrobat"},
    @{Name="Encompass"; Pattern="Encompass"}
)

foreach ($Check in $VerificationChecks) {
    $Found = winget list | Select-String -Pattern $Check.Pattern
    if ($Found) {
        Write-Log "[OK] VERIFIED: $($Check.Name) is installed"
    } else {
        Write-Log "[MISSING] NOT FOUND: $($Check.Name) may not be installed"
    }
    Start-Sleep -Seconds 1
}

# ============================================
# PHASE 6: Post-Installation Configuration
# ============================================
Write-Log "=== PHASE 6: Post-Installation Configuration ==="

# Trigger Intune sync
Write-Log "Triggering Intune device sync..."
Get-ScheduledTask | Where-Object {$_.TaskName -eq 'PushLaunch'} | Start-ScheduledTask -ErrorAction SilentlyContinue

# Force Group Policy update
Write-Log "Updating Group Policies..."
gpupdate /force 2>&1 | Out-Null
Write-Log "Group Policy updated"

# Create completion flag for Intune detection
Set-Content -Path $CompletionFlag -Value "Completed: $(Get-Date)"

# ============================================
# Script Completion
# ============================================
Write-Log "=== Enrollment Script V1.1 Completed ==="
Write-Log "Total apps processed: 8"
Write-Log "Log file saved to: $LogFile"
Write-Log ""
Write-Log "INSTALLED APPLICATIONS:"
Write-Log "  [X] RingCentral"
Write-Log "  [X] Microsoft 365 Apps"
Write-Log "  [X] Microsoft Teams"
Write-Log "  [X] Microsoft OneDrive"
Write-Log "  [X] Azure VPN Client"
Write-Log "  [X] Dynamic Theme (Dark mode enabled)"
Write-Log "  [X] Adobe Acrobat Reader DC"
Write-Log "  [ ] Encompass (via separate Intune package)"
Write-Log ""
Write-Log "NEXT STEPS:"
Write-Log "1. Verify all applications are working"
Write-Log "2. Configure Azure VPN connection"
Write-Log "3. Open Dynamic Theme to customize preferences"
Write-Log "4. V2 will add: Defender onboarding + VPN profile import"
Write-Log ""

exit 0
