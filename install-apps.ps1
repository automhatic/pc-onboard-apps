# ============================================================================
# Windows Enrollment Script V1.1
# Installs standard applications via Winget for new PC onboarding
# ============================================================================

# Create log directory
$LogDir = "C:\ProgramData\EnrollmentScript"
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path $LogDir "EnrollmentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

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
$Winget = Get-Command winget -ErrorAction SilentlyContinue
if (-not $Winget) {
    Write-Log "ERROR: Winget not found. Exiting."
    exit 1
}
Write-Log "Winget found at: $($Winget.Source)"

# Application list
$Apps = @(
    @{ Name = "RingCentral"; ID = "RingCentral.RingCentral" },
    @{ Name = "Microsoft 365 Apps"; ID = "Microsoft.Office" },
    @{ Name = "Microsoft Teams"; ID = "Microsoft.Teams" },
    @{ Name = "Microsoft OneDrive"; ID = "Microsoft.OneDrive" },
    @{ Name = "Azure VPN Client"; ID = "9NP355QT2SQB" },
    @{ Name = "Dynamic Theme"; ID = "9NBLGGH1ZBKW" }
)

# === PHASE 1: Installing Standard Applications ===
Write-Log "=== PHASE 1: Installing Standard Applications ==="

$AppCount = 0
foreach ($App in $Apps) {
    $AppCount++
    Write-Log "Installing: $($App.Name) ($($App.ID))"
    
    try {
        # Check if already installed
        $Installed = winget list --id $App.ID --exact 2>&1
        
        if ($Installed -match $App.ID) {
            Write-Log "INFO: $($App.Name) already installed, skipping"
            Start-Sleep -Seconds 3
            continue
        }
        
        # Install the app
        $InstallArgs = @(
            "install",
            "--id", $App.ID,
            "--exact",
            "--silent",
            "--accept-source-agreements",
            "--accept-package-agreements"
        )
        
        & winget $InstallArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "SUCCESS: $($App.Name) installed"
        } else {
            Write-Log "WARNING: $($App.Name) installation returned code $LASTEXITCODE"
        }
        
    } catch {
        Write-Log "ERROR: Failed to install $($App.Name): $($_.Exception.Message)"
    }
    
    Start-Sleep -Seconds 3
}

# === PHASE 2: Installing Adobe Acrobat Reader ===
Write-Log "=== PHASE 2: Installing Adobe Acrobat Reader ==="

try {
    Write-Log "Installing Adobe Acrobat Reader (64-bit, no add-ons)..."
    
    $AdobeInstall = winget install --id Adobe.Acrobat.Reader.64-bit `
        --exact `
        --silent `
        --override "/sPB /rs /msi EULA_ACCEPT=YES" `
        --accept-source-agreements `
        --accept-package-agreements 2>&1
    
    Write-Log "$AdobeInstall"
    
    if ($LASTEXITCODE -eq 0) {
        Write-Log "SUCCESS: Adobe Reader installed"
    } else {
        Write-Log "WARNING: Adobe Reader installation returned code $LASTEXITCODE"
    }
    
    Write-Log "Waiting for Adobe Reader installation to complete..."
    Start-Sleep -Seconds 15
    
} catch {
    Write-Log "ERROR: Adobe Reader installation failed: $($_.Exception.Message)"
}

# === PHASE 3: Installing Encompass Smart Client ===
Write-Log "=== PHASE 3: Installing Encompass Smart Client ==="
Start-Sleep -Seconds 3

# Check if Adobe Reader is installed (Encompass dependency)
$AdobeInstalled = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $_.DisplayName -like "*Adobe Acrobat*" }

if (-not $AdobeInstalled) {
    Write-Log "WARNING: Adobe Reader not detected in registry, but continuing anyway..."
}

# Check for Encompass installer
$EncompassInstallerPath = "C:\Temp\EncompassInstaller.exe"

if (Test-Path $EncompassInstallerPath) {
    Write-Log "Found Encompass installer at: $EncompassInstallerPath"
    Write-Log "Installing Encompass Smart Client..."
    
    try {
        $EncompassInstall = Start-Process -FilePath $EncompassInstallerPath `
            -ArgumentList "/S" `
            -Wait `
            -PassThru
        
        if ($EncompassInstall.ExitCode -eq 0) {
            Write-Log "SUCCESS: Encompass Smart Client installed"
        } else {
            Write-Log "WARNING: Encompass installation returned code $($EncompassInstall.ExitCode)"
        }
        
    } catch {
        Write-Log "ERROR: Encompass installation failed: $($_.Exception.Message)"
    }
    
} else {
    Write-Log "INFO: Encompass not found - will be installed via separate Intune package"
}

Start-Sleep -Seconds 3

# === PHASE 4: Installation Verification ===
Write-Log "=== PHASE 4: Installation Verification ==="
Start-Sleep -Seconds 3

$VerificationChecks = @(
    @{ Name = "RingCentral"; Pattern = "RingCentral" },
    @{ Name = "Microsoft 365"; Pattern = "Microsoft 365" },
    @{ Name = "Teams"; Pattern = "Microsoft Teams" },
    @{ Name = "OneDrive"; Pattern = "OneDrive" },
    @{ Name = "Azure VPN"; Pattern = "Azure VPN" },
    @{ Name = "Dynamic Theme"; Pattern = "Dynamic Theme" },
    @{ Name = "Adobe"; Pattern = "Adobe Acrobat" },
    @{ Name = "Encompass"; Pattern = "Encompass" }
)

foreach ($Check in $VerificationChecks) {
    try {
        $Found = winget list | Select-String -Pattern $Check.Pattern
        
        if ($Found) {
            Write-Log "[OK] VERIFIED: $($Check.Name) is installed"
        } else {
            Write-Log "[MISSING] NOT FOUND: $($Check.Name) may not be installed"
        }
        
        Start-Sleep -Seconds 1
        
    } catch {
        Write-Log "[ERROR] Could not verify $($Check.Name): $($_.Exception.Message)"
    }
}

# === PHASE 5: Post-Installation Configuration ===
Write-Log "=== PHASE 5: Post-Installation Configuration ==="

# Trigger Intune sync
Write-Log "Triggering Intune device sync..."
try {
    $IntuneSync = Get-ScheduledTask | Where-Object { $_.TaskName -eq "PushLaunch" }
    if ($IntuneSync) {
        Start-ScheduledTask -TaskName "PushLaunch" -ErrorAction SilentlyContinue
    }
} catch {
    Write-Log "WARNING: Could not trigger Intune sync: $($_.Exception.Message)"
}

# Update Group Policy
Write-Log "Updating Group Policies..."
try {
    gpupdate /force | Out-Null
    Write-Log "Group Policy updated"
} catch {
    Write-Log "WARNING: Group Policy update failed: $($_.Exception.Message)"
}

# Final summary
Write-Log "=== Enrollment Script V1.1 Completed ==="
Write-Log "Total apps processed: $AppCount"

Start-Sleep -Seconds 1

Write-Log "Log file saved to: $LogFile"
Write-Log ""
Write-Log "INSTALLED APPLICATIONS:"
Write-Log "  [X] RingCentral"
Write-Log "  [X] Microsoft 365 Apps"
Write-Log "  [X] Microsoft Teams"
Write-Log "  [X] Microsoft OneDrive"
Write-Log "  [X] Azure VPN Client"
Write-Log "  [X] Dynamic Theme"
Write-Log "  [X] Adobe Acrobat Reader DC"
Write-Log "  [ ] Encompass (via separate Intune package)"
Write-Log ""
Write-Log "NEXT STEPS:"
Write-Log "1. Verify all applications are working"
Write-Log "2. Configure Azure VPN connection"
Write-Log "3. Open Dynamic Theme to set your preferred light/dark mode"
Write-Log "4. V2 will add: Defender onboarding + VPN profile import"
Write-Log ""

exit 0
