<#
.SYNOPSIS
    Enrollment Script V1 - App Installation
.DESCRIPTION
    Installs: RingCentral, M365 Apps, Teams, OneDrive, Adobe Reader, Azure VPN Client, Encompass
    Handles admin requirements and installation order dependencies
.NOTES
    Deploy via Intune as Win32 app - Run as SYSTEM context
    Version: 1.0
#>

# Set up logging
$LogPath = "C:\ProgramData\EnrollmentScript"
$LogFile = "$LogPath\EnrollmentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

if (!(Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp - $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

Write-Log "=== Enrollment Script V1 Started ==="
Write-Log "Running as: $env:USERNAME"

# ============================================
# Verify Winget is Available
# ============================================
Write-Log "Verifying Winget availability..."

try {
    $wingetPath = (Get-Command winget -ErrorAction Stop).Source
    Write-Log "Winget found at: $wingetPath"
} catch {
    Write-Log "ERROR: Winget not found. Attempting to install..."
    try {
        # For Windows 10/11 - Install App Installer
        Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe
        Start-Sleep -Seconds 10
        Write-Log "Winget installed successfully"
    } catch {
        Write-Log "CRITICAL ERROR: Cannot install Winget. Exiting."
        exit 1
    }
}

# ============================================
# PHASE 1: Install Standard Apps
# ============================================
Write-Log "=== PHASE 1: Installing Standard Applications ==="

# App definitions with Winget IDs
$StandardApps = @(
    @{Name="RingCentral"; ID="RingCentral.RingCentral"},
    @{Name="Microsoft 365 Apps"; ID="Microsoft.Office"},
    @{Name="Microsoft Teams"; ID="Microsoft.Teams"},
    @{Name="Microsoft OneDrive"; ID="Microsoft.OneDrive"},
    @{Name="Azure VPN Client"; ID="Microsoft.AzureVPNClient"}
)

foreach ($App in $StandardApps) {
    Write-Log "Installing: $($App.Name) ($($App.ID))"
    try {
        # Check if already installed
        $checkInstall = winget list --id $App.ID --exact 2>&1
        if ($checkInstall -match $App.ID) {
            Write-Log "INFO: $($App.Name) already installed, skipping..."
            continue
        }

        # Install the app
        $result = winget install --id $App.ID --silent --accept-package-agreements --accept-source-agreements --scope machine 2>&1
        
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            # -1978335189 = already installed (hex: 0x8A15000B)
            Write-Log "SUCCESS: $($App.Name) installed"
        } else {
            Write-Log "WARNING: $($App.Name) installation returned code $LASTEXITCODE"
        }
        
        # Small delay between installations
        Start-Sleep -Seconds 3
        
    } catch {
        Write-Log "ERROR installing $($App.Name): $_"
    }
}

# ============================================
# PHASE 2: Install Adobe Reader (No Bloatware)
# ============================================
Write-Log "=== PHASE 2: Installing Adobe Acrobat Reader ==="

try {
    # Check if already installed
    $adobeInstalled = winget list --id Adobe.Acrobat.Reader.64-bit --exact 2>&1
    
    if ($adobeInstalled -match "Adobe.Acrobat.Reader") {
        Write-Log "INFO: Adobe Reader already installed"
    } else {
        Write-Log "Installing Adobe Acrobat Reader (64-bit, no add-ons)..."
        
        # Winget install with override to skip McAfee and other bloat
        # The /msi parameter passes arguments to the MSI installer
        $result = winget install --id Adobe.Acrobat.Reader.64-bit --silent --accept-package-agreements --accept-source-agreements --override "/quiet /norestart EULA_ACCEPT=YES DISABLE_ARM_SERVICE_INSTALL=1" 2>&1
        
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            Write-Log "SUCCESS: Adobe Reader installed"
        } else {
            Write-Log "WARNING: Adobe Reader installation returned code $LASTEXITCODE"
        }
        
        # Wait for Adobe to fully install before proceeding to Encompass
        Write-Log "Waiting for Adobe Reader installation to complete..."
        Start-Sleep -Seconds 15
    }
} catch {
    Write-Log "ERROR installing Adobe Reader: $_"
}

# ============================================
# PHASE 3: Install Encompass (Requires Admin + Adobe)
# ============================================
Write-Log "=== PHASE 3: Installing Encompass Smart Client ==="

try {
    # Verify Adobe is installed first
    $adobeCheck = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                  Where-Object { $_.DisplayName -like "*Adobe*Reader*" }
    
    if (-not $adobeCheck) {
        $adobeCheck = Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                      Where-Object { $_.DisplayName -like "*Adobe*Reader*" }
    }

    if ($adobeCheck) {
        Write-Log "Adobe Reader verified - proceeding with Encompass installation"
    } else {
        Write-Log "WARNING: Adobe Reader not detected in registry, but continuing anyway..."
    }

    # Check if already installed
    $encompassInstalled = winget list --id ICEMortgage.Encompass --exact 2>&1
    
    if ($encompassInstalled -match "ICEMortgage.Encompass") {
        Write-Log "INFO: Encompass already installed"
    } else {
        Write-Log "Installing Encompass Smart Client (requires admin rights)..."
        
        # Install with elevated privileges (script should already be running as SYSTEM)
        $result = winget install --id ICEMortgage.Encompass --version 2.0.1.0 --silent --accept-package-agreements --accept-source-agreements --scope machine 2>&1
        
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            Write-Log "SUCCESS: Encompass Smart Client installed"
        } else {
            Write-Log "WARNING: Encompass installation returned code $LASTEXITCODE"
            Write-Log "Output: $result"
        }
    }
} catch {
    Write-Log "ERROR installing Encompass: $_"
}

# ============================================
# PHASE 4: Verification
# ============================================
Write-Log "=== PHASE 4: Installation Verification ==="

$AppsToVerify = @(
    "RingCentral",
    "Microsoft 365",
    "Teams",
    "OneDrive",
    "Adobe",
    "Azure VPN",
    "Encompass"
)

foreach ($AppName in $AppsToVerify) {
    $installed = winget list | Select-String -Pattern $AppName -Quiet
    if ($installed) {
        Write-Log "[OK] VERIFIED: $AppName is installed"
    } else {
        Write-Log "[MISSING] NOT FOUND: $AppName may not be installed"
    }
}

# ============================================
# PHASE 5: Post-Installation Tasks
# ============================================
Write-Log "=== PHASE 5: Post-Installation Configuration ==="

# Trigger Intune sync
Write-Log "Triggering Intune device sync..."
try {
    $IntuneAgent = "$env:ProgramFiles\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe"
    if (Test-Path $IntuneAgent) {
        Start-Process -FilePath $IntuneAgent -ArgumentList "intunemanagementextension://syncapp" -NoNewWindow -ErrorAction SilentlyContinue
        Write-Log "Intune sync triggered"
    }
} catch {
    Write-Log "Could not trigger Intune sync: $_"
}

# Force Group Policy update
Write-Log "Updating Group Policies..."
try {
    gpupdate /force | Out-Null
    Write-Log "Group Policy updated"
} catch {
    Write-Log "Group Policy update failed: $_"
}

# ============================================
# Completion
# ============================================
Write-Log "=== Enrollment Script V1 Completed ==="
Write-Log "Total apps processed: 7"
Write-Log "Log file saved to: $LogFile"
Write-Log ""
Write-Log "NEXT STEPS:"
Write-Log "1. Verify all applications are working"
Write-Log "2. V2 will add: Defender onboarding + VPN profile import"
Write-Log ""

# Create a completion flag file
$CompletionFlag = "$LogPath\enrollment_v1_complete.txt"
Set-Content -Path $CompletionFlag -Value "Completed: $(Get-Date)"

# Return success
exit 0
