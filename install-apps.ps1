<#
.SYNOPSIS
    Windows Device Enrollment Script V1.3 - System Context Apps Only
.DESCRIPTION
    Automatically installs system-context applications for new Windows devices
    Designed to run via Intune as SYSTEM account during Autopilot enrollment
    User-context apps (RingCentral, Encompass) are deployed separately via Intune
.NOTES
    Author: IT Department
    Version: 1.3
    Last Updated: 2026-04-19
    - Removed user-context apps (RingCentral, Encompass)
    - Removed Store apps (Dynamic Theme)
    - Focused on system-context apps only
    - Added Intune deployment notes
#>

#Requires -RunAsAdministrator

# ============================================
# CONFIGURATION
# ============================================

$LogFolder = "C:\ProgramData\EnrollmentScript"
$LogFile = Join-Path $LogFolder "EnrollmentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# System-Context Application List
$StandardApps = @(
    @{ Name = "Microsoft 365 Apps"; ID = "Microsoft.Office" }
    @{ Name = "Microsoft Teams"; ID = "Microsoft.Teams" }
    @{ Name = "Microsoft OneDrive"; ID = "Microsoft.OneDrive" }
    @{ Name = "Azure VPN Client"; ID = "9NP355QT2SQB" }
)

# ============================================
# FUNCTIONS
# ============================================

function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$Timestamp - $Message"
    Add-Content -Path $LogFile -Value $LogMessage -Force
    Write-Host $LogMessage
}

function Write-Progress-Bar {
    param(
        [string]$Activity,
        [int]$PercentComplete,
        [string]$Status
    )
    
    $barLength = 40
    $filled = [math]::Floor($barLength * $PercentComplete / 100)
    $empty = $barLength - $filled
    
    $bar = "[" + ("█" * $filled) + ("░" * $empty) + "]"
    
    Write-Host "`r$Activity $bar $PercentComplete% - $Status" -NoNewline -ForegroundColor Cyan
}

function Test-WingetAvailable {
    Write-Log "Verifying Winget availability..."
    Write-Host "  → Searching for winget..." -ForegroundColor Gray
    
    try {
        $WingetPath = (Get-Command winget -ErrorAction Stop).Source
        Write-Log "Winget found at: $WingetPath"
        Write-Host "  ✓ Winget found: $WingetPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Log "ERROR: Winget not found!"
        Write-Host "  ✗ Winget not found!" -ForegroundColor Red
        return $false
    }
}

function Install-App {
    param($App, $CurrentNum, $TotalApps)
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  APP $CurrentNum of $TotalApps: $($App.Name)" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Log "Installing: $($App.Name) ($($App.ID))"
    
    try {
        # Step 1: Check if already installed
        Write-Host "  [1/4] Checking if already installed..." -ForegroundColor Yellow
        $Installed = winget list --id $App.ID --exact 2>&1
        
        if ($Installed -match $App.ID) {
            Write-Log "  ✓ $($App.Name) is already installed"
            Write-Host "  ✓ Already installed - Skipping" -ForegroundColor Green
            return $true
        }
        
        Write-Host "  → Not installed, proceeding with installation" -ForegroundColor Gray
        
        # Step 2: Download
        Write-Host "  [2/4] Downloading $($App.Name)..." -ForegroundColor Yellow
        Write-Host "  → Contacting winget repository..." -ForegroundColor Gray
        
        # Step 3: Install
        Write-Host "  [3/4] Installing (this may take 2-10 minutes)..." -ForegroundColor Yellow
        Write-Host "  → Running silent installation..." -ForegroundColor Gray
        
        $InstallArgs = @(
            "install"
            "--id", $App.ID
            "--exact"
            "--silent"
            "--accept-package-agreements"
            "--accept-source-agreements"
            "--disable-interactivity"
        )
        
        # Create temp files for output
        $outFile = "$env:TEMP\winget-out-$($App.ID).txt"
        $errFile = "$env:TEMP\winget-err-$($App.ID).txt"
        
        # Start the process
        $process = Start-Process -FilePath "winget" `
            -ArgumentList $InstallArgs `
            -NoNewWindow `
            -PassThru `
            -RedirectStandardOutput $outFile `
            -RedirectStandardError $errFile
        
        # Monitor progress
        $timeout = 600 # 10 minutes
        $elapsed = 0
        $lastCheck = Get-Date
        
        Write-Host "  → Installation in progress" -NoNewline -ForegroundColor Gray
        
        while (-not $process.HasExited -and $elapsed -lt $timeout) {
            Start-Sleep -Seconds 3
            $elapsed += 3
            
            # Show activity indicator
            Write-Host "." -NoNewline -ForegroundColor Gray
            
            # Check for installer processes every 15 seconds
            if ($elapsed % 15 -eq 0) {
                $installerNames = @("msiexec", "setup", "install", $App.Name.Split(" ")[0])
                $activeInstaller = Get-Process | Where-Object { 
                    $installerNames | ForEach-Object { 
                        if ($_.ProcessName -match $_) { return $true }
                    }
                }
                
                if ($activeInstaller) {
                    Write-Host "!" -NoNewline -ForegroundColor Yellow # Installer detected
                }
            }
            
            # Progress update every 30 seconds
            if ($elapsed % 30 -eq 0) {
                $minutes = [math]::Floor($elapsed / 60)
                $seconds = $elapsed % 60
                Write-Host " [{0:D2}:{1:D2}]" -f $minutes, $seconds -NoNewline -ForegroundColor DarkGray
            }
        }
        
        Write-Host "" # New line after progress dots
        
        # Check if timed out
        if ($elapsed -ge $timeout) {
            Write-Log "  ⚠ $($App.Name) installation timed out after $timeout seconds"
            Write-Host "  ⚠ Installation timed out - may still be running in background" -ForegroundColor Yellow
            try { $process.Kill() } catch {}
            return $false
        }
        
        $process.WaitForExit()
        
        # Step 4: Verify
        Write-Host "  [4/4] Verifying installation..." -ForegroundColor Yellow
        
        # Read output
        $output = ""
        $errorOutput = ""
        
        if (Test-Path $outFile) {
            $output = Get-Content $outFile -Raw -ErrorAction SilentlyContinue
        }
        if (Test-Path $errFile) {
            $errorOutput = Get-Content $errFile -Raw -ErrorAction SilentlyContinue
        }
        
        # Check exit code
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Log "  ✓ $($App.Name) installed successfully (Exit code: $($process.ExitCode))"
            Write-Host "  ✓ Installation completed successfully!" -ForegroundColor Green
            
            if ($process.ExitCode -eq 3010) {
                Write-Host "  ℹ Reboot required for this application" -ForegroundColor Cyan
            }
            
            return $true
            
        } elseif ($process.ExitCode -eq -1978335189) {
            # Already installed (different version)
            Write-Log "  ✓ $($App.Name) already installed (different version)"
            Write-Host "  ✓ Already installed (different version)" -ForegroundColor Green
            return $true
            
        } else {
            Write-Log "  ✗ $($App.Name) installation failed (Exit code: $($process.ExitCode))"
            Write-Log "  Output: $output"
            Write-Log "  Error: $errorOutput"
            Write-Host "  ✗ Installation failed (Exit code: $($process.ExitCode))" -ForegroundColor Red
            
            if ($errorOutput) {
                Write-Host "  Error details: $($errorOutput.Substring(0, [Math]::Min(200, $errorOutput.Length)))" -ForegroundColor Red
            }
            
            return $false
        }
        
    } catch {
        Write-Log "ERROR: Failed to install $($App.Name): $($_.Exception.Message)"
        Write-Host "  ✗ Exception: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    } finally {
        # Cleanup temp files
        Remove-Item "$env:TEMP\winget-out-$($App.ID).txt" -ErrorAction SilentlyContinue
        Remove-Item "$env:TEMP\winget-err-$($App.ID).txt" -ErrorAction SilentlyContinue
    }
}

function Install-AdobeReader {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  ADOBE ACROBAT READER" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Log "Installing Adobe Acrobat Reader (64-bit, no add-ons)..."
    
    try {
        Write-Host "  [1/3] Checking if already installed..." -ForegroundColor Yellow
        $AdobeInstalled = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" |
            Where-Object { $_.DisplayName -match "Adobe Acrobat Reader" }
        
        if ($AdobeInstalled) {
            Write-Log "  ✓ Adobe Reader already installed: $($AdobeInstalled.DisplayName)"
            Write-Host "  ✓ Already installed - Skipping" -ForegroundColor Green
            return $true
        }
        
        Write-Host "  [2/3] Downloading and installing..." -ForegroundColor Yellow
        Write-Host "  → This may take 3-5 minutes..." -ForegroundColor Gray
        
        $AdobeInstall = winget install --id Adobe.Acrobat.Reader.64-bit `
            --exact --silent --accept-package-agreements --accept-source-agreements `
            --override "/sPB /rs /msi" 2>&1
        
        Write-Host "  [3/3] Verifying installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        
        $AdobeCheck = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" |
            Where-Object { $_.DisplayName -match "Adobe Acrobat Reader" }
        
        if ($AdobeCheck) {
            Write-Log "  ✓ Adobe Reader installed successfully"
            Write-Host "  ✓ Installation completed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Log "  ⚠ Adobe Reader installation completed but not detected in registry"
            Write-Host "  ⚠ Installation may need verification" -ForegroundColor Yellow
            return $false
        }
        
    } catch {
        Write-Log "ERROR: Adobe Reader installation failed: $($_.Exception.Message)"
        Write-Host "  ✗ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Show-IntuneDeploymentInfo {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  ADDITIONAL APPS VIA INTUNE" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Log "Displaying Intune deployment information..."
    
    Write-Host ""
    Write-Host "  The following apps will be deployed separately via Intune:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  USER-CONTEXT APPS (Win32):" -ForegroundColor Cyan
    Write-Host "    • RingCentral" -ForegroundColor White
    Write-Host "      → Installs to user profile after sign-in" -ForegroundColor Gray
    Write-Host "      → Assignment: Required" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    • Encompass Smart Client" -ForegroundColor White
    Write-Host "      → Installs to user profile after sign-in" -ForegroundColor Gray
    Write-Host "      → Assignment: Required (when configured)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  OPTIONAL APPS (Microsoft Store):" -ForegroundColor Cyan
    Write-Host "    • Dynamic Theme" -ForegroundColor White
    Write-Host "      → Available in Company Portal" -ForegroundColor Gray
    Write-Host "      → Assignment: Available for enrolled devices" -ForegroundColor Gray
    Write-Host ""
    
    Write-Log "INFO: User-context apps will install automatically after user signs in"
    Write-Log "INFO: Optional Store apps available in Company Portal"
    
    Write-Host "  ℹ These apps install in USER context and cannot be installed" -ForegroundColor Cyan
    Write-Host "    during SYSTEM-context enrollment" -ForegroundColor Cyan
    Write-Host ""
}

function Invoke-PostConfiguration {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  POST-INSTALLATION CONFIGURATION" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Log "Starting post-installation configuration..."
    
    # Intune Sync
    Write-Host "  [1/2] Triggering Intune device sync..." -ForegroundColor Yellow
    Write-Host "  → Contacting Intune service..." -ForegroundColor Gray
    try {
        $IntuneAgentPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe"
        if (Test-Path $IntuneAgentPath) {
            Start-Process -FilePath $IntuneAgentPath -ArgumentList "-SyncNow" -NoNewWindow -ErrorAction SilentlyContinue
            Write-Log "Intune sync triggered"
            Write-Host "  ✓ Intune sync triggered (user-context apps will install after sign-in)" -ForegroundColor Green
        } else {
            Write-Log "INFO: Intune Management Extension not found (device may not be enrolled yet)"
            Write-Host "  ℹ Intune not detected (will sync after enrollment completes)" -ForegroundColor Cyan
        }
    } catch {
        Write-Log "WARNING: Could not trigger Intune sync: $($_.Exception.Message)"
        Write-Host "  ⚠ Could not trigger Intune sync (will sync automatically)" -ForegroundColor Yellow
    }
    
    # Group Policy Update
    Write-Host "  [2/2] Updating Group Policies..." -ForegroundColor Yellow
    Write-Host "  → Running gpupdate..." -ForegroundColor Gray
    try {
        Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -NoNewWindow -Wait
        Write-Log "Group Policy updated"
        Write-Host "  ✓ Group Policy updated" -ForegroundColor Green
    } catch {
        Write-Log "WARNING: Could not update Group Policy: $($_.Exception.Message)"
        Write-Host "  ⚠ Could not update Group Policy" -ForegroundColor Yellow
    }
}

# ============================================
# MAIN SCRIPT
# ============================================

# Create log folder
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

# Script header
Clear-Host
Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                                                            ║" -ForegroundColor Cyan
Write-Host "║        WINDOWS DEVICE ENROLLMENT SCRIPT V1.3               ║" -ForegroundColor White
Write-Host "║        System Context Apps Only                            ║" -ForegroundColor Gray
Write-Host "║                                                            ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

Write-Log "=== Enrollment Script V1.3 Started ==="
Write-Log "Running as: $env:COMPUTERNAME\$env:USERNAME"

Write-Host "  Computer: $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  User Context: $env:USERNAME" -ForegroundColor White
Write-Host "  Log File: $LogFile" -ForegroundColor Gray
Write-Host ""

# Verify Winget
if (-not (Test-WingetAvailable)) {
    Write-Log "FATAL: Cannot proceed without Winget"
    Write-Host ""
    Write-Host "  ✗ FATAL ERROR: Winget is required but not found" -ForegroundColor Red
    Write-Host "  Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host ""
Start-Sleep -Seconds 2

# ============================================
# PHASE 1: System-Context Applications
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  PHASE 1: INSTALLING SYSTEM-CONTEXT APPLICATIONS          ║" -ForegroundColor Green
Write-Host "║  Total Apps: $($StandardApps.Count)                                                   ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Log "=== PHASE 1: Installing System-Context Applications ==="

$SuccessCount = 0
$FailCount = 0
$AppNum = 1

foreach ($App in $StandardApps) {
    $Result = Install-App -App $App -CurrentNum $AppNum -TotalApps $StandardApps.Count
    
    if ($Result) {
        $SuccessCount++
    } else {
        $FailCount++
    }
    
    $AppNum++
    Start-Sleep -Seconds 2
}

Write-Host ""
Write-Host "  Phase 1 Summary:" -ForegroundColor Cyan
Write-Host "  ✓ Successful: $SuccessCount" -ForegroundColor Green
if ($FailCount -gt 0) {
    Write-Host "  ✗ Failed: $FailCount" -ForegroundColor Red
}
Write-Host ""

# ============================================
# PHASE 2: Adobe Acrobat Reader
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  PHASE 2: INSTALLING ADOBE ACROBAT READER                 ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Log "=== PHASE 2: Installing Adobe Acrobat Reader ==="

$AdobeResult = Install-AdobeReader
if ($AdobeResult) {
    $SuccessCount++
} else {
    $FailCount++
}
Start-Sleep -Seconds 2

# ============================================
# PHASE 3: Intune Deployment Information
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  PHASE 3: INTUNE-DEPLOYED APPLICATIONS                    ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Log "=== PHASE 3: Intune Deployment Information ==="

Show-IntuneDeploymentInfo
Start-Sleep -Seconds 2

# ============================================
# PHASE 4: Installation Verification
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  PHASE 4: INSTALLATION VERIFICATION                       ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Log "=== PHASE 4: Installation Verification ==="

Write-Host ""
Write-Host "  Verifying system-context applications..." -ForegroundColor Yellow
Write-Host ""

$VerificationChecks = @(
    @{ Name = "Microsoft 365"; Pattern = "Microsoft 365|Office" }
    @{ Name = "Teams"; Pattern = "Teams" }
    @{ Name = "OneDrive"; Pattern = "OneDrive" }
    @{ Name = "Azure VPN"; Pattern = "Azure VPN" }
    @{ Name = "Adobe Reader"; Pattern = "Adobe Acrobat" }
)

$InstalledCount = 0
$NotInstalledCount = 0

foreach ($Check in $VerificationChecks) {
    Write-Host "  Checking $($Check.Name)..." -NoNewline -ForegroundColor Gray
    
    try {
        $Found = winget list | Select-String -Pattern $Check.Pattern
        
        if ($Found) {
            Write-Host " ✓" -ForegroundColor Green
            Write-Log "[VERIFIED] $($Check.Name) is installed"
            $InstalledCount++
        } else {
            Write-Host " ✗" -ForegroundColor Red
            Write-Log "[MISSING] $($Check.Name) not found"
            $NotInstalledCount++
        }
    } catch {
        Write-Host " ?" -ForegroundColor Yellow
        Write-Log "[ERROR] Could not verify $($Check.Name): $($_.Exception.Message)"
        $NotInstalledCount++
    }
    
    Start-Sleep -Milliseconds 500
}

Write-Host ""
Write-Host "  Verification Summary:" -ForegroundColor Cyan
Write-Host "  ✓ Installed: $InstalledCount" -ForegroundColor Green
if ($NotInstalledCount -gt 0) {
    Write-Host "  ✗ Not Installed: $NotInstalledCount" -ForegroundColor Red
}
Write-Host ""

# ============================================
# PHASE 5: Post-Configuration
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  PHASE 5: POST-INSTALLATION CONFIGURATION                 ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Log "=== PHASE 5: Post-Installation Configuration ==="

Invoke-PostConfiguration

# ============================================
# COMPLETION
# ============================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                                                            ║" -ForegroundColor Green
Write-Host "║  ✓ ENROLLMENT SCRIPT COMPLETED SUCCESSFULLY!               ║" -ForegroundColor White
Write-Host "║                                                            ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""

Write-Log "=== Enrollment Script V1.3 Completed ==="
Write-Log "System-context apps installed: $SuccessCount | Failed: $FailCount"

Write-Host "  Final Summary:" -ForegroundColor Cyan
Write-Host "  • System-Context Apps Installed: $SuccessCount" -ForegroundColor White
Write-Host "  • System-Context Apps Failed: $FailCount" -ForegroundColor White
Write-Host "  • Verified Installations: $InstalledCount" -ForegroundColor White
Write-Host "  • User-Context Apps: Will install via Intune after sign-in" -ForegroundColor Cyan
Write-Host "  • Log File: $LogFile" -ForegroundColor Gray
Write-Host ""

if ($FailCount -gt 0 -or $NotInstalledCount -gt 0) {
    Write-Host "  ⚠ Some installations may need attention" -ForegroundColor Yellow
    Write-Host "  → Review the log file for details" -ForegroundColor Yellow
} else {
    Write-Host "  ✓ All system-context apps completed successfully!" -ForegroundColor Green
    Write-Host "  ✓ User-context apps will install automatically after user signs in" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Press any key to close..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

exit 0
