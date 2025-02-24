# Set global variables
$hostname = hostname

if ($hostname.Substring(0, 4) -match '^\d{4}') {
    $siteNumber = $hostname.Substring(0, 4)
} else {
    Write-Error "This computer, '${hostname}', is either named incorrectly of is not a back office workstation. Exiting script..."
    Start-Sleep -Seconds 10
    Exit
}

$logFileName = "${hostname}_Backup.log"
$scriptName = [System.IO.Path]::GetFileName($MyInvocation.ScriptName)

# Function to log messages
function LogMessage {
    param (
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try {
        Add-Content -Path $logFileName -Value "$TimeStamp - $Message"
    } catch {
        Write-Error "$scriptName failed to write $Message to $logFileName on $hostname. Error: $_"
    }
}

# Function to handle network interruptions during file operations
function HandleNetworkInterruptions {
    param (
        [ScriptBlock]$Operation
    )
    $retryCount = 3
    for ($i = 1; $i -le $retryCount; $i++) {
        try {
            & $Operation
            LogMessage "Operation succeeded on attempt $i."
            return $true
        } catch [System.IO.IOException] {
            LogMessage "Network interruption occurred. Retrying operation ($i/$retryCount)..."
            Start-Sleep -Seconds 5
        }
    }
    LogMessage "Operation failed after $retryCount retries due to network interruptions."
    return $false
}

# Function to check for backup directory corruption
function CheckBackupDirCorruption {
    param (
        [string]$BackupDir
    )
    if (!(Test-Path -Path $BackupDir -ErrorAction SilentlyContinue)) {
        LogMessage "$'{BackupDir}' doesn't exist, assumed corrupted" "ERROR"
        return $true
    }
    # Check for actual corruption or inaccessibility
    try {
        Get-ChildItem $BackupDir -ErrorAction Stop | Out-Null
        LogMessage "'${BackupDir}' is intact and accessible"
        return $false
    } catch {
        LogMessage "'${BackupDir}' is corrupted or inaccessible" "ERROR"
        return $true
    }
}

# Function to reset permissions
function ResetPermissions {
    param (
        [string]$User,
        [string]$Path
    )
    try {
        if (Test-Path -Path $Path) {
            icacls $Path /reset /T | Out-Null
            LogMessage "${User}'s permissions have been reset for: $Path"
            return $true  # Permission reset succeeded
        } else {
            LogMessage "Failed to reset permissions for ${User}, path does not exist: $Path" "ERROR"
            return $false  # Permission reset failed
        }
    } catch {
        LogMessage "An error occurred while resetting ${User}'s permissions for: $Path. Error: $_"
        return $false  # Permission reset failed
    }
}

# Function to restore browser settings
function RestoreBrowserSettings {
    param (
        [string]$User,
        [string]$ProfileDir,
        [string]$BackupDir,
        [string[]]$FilesToRestore,
        [string[]]$FoldersToRestore
    )
    try {
        if (Test-Path -Path $ProfileDir) {
            foreach ($file in $FilesToRestore) {
                $backupPath = Join-Path -Path $BackupDir -ChildPath $file
                if (Test-Path -Path $backupPath) {
                    if (HandleNetworkInterruptions {
                        Copy-Item -Path $backupPath -Destination $ProfileDir -Recurse -Force
                    }) {
                        LogMessage "Restored file for ${User}: $file"
                    } else {
                        throw "Unstable network."
                    }
                }
            }
            foreach ($folder in $FoldersToRestore) {
                $backupPath = Join-Path -Path $BackupDir -ChildPath $folder
                if (Test-Path -Path $backupPath) {
                    $destinationPath = Join-Path -Path $ProfileDir -ChildPath $folder
                    if (HandleNetworkInterruptions {
                        Copy-Item -Path $backupPath -Destination $destinationPath -Recurse -Force
                    }) {
                        LogMessage "Restored folder for ${User}: $folder"
                    } else {
                        throw "Unstable network."
                    }
                }
            }
            return $true  # Restoration succeeded
        } else {
            LogMessage "Profile directory does not exist for ${User}: $ProfileDir"
            return $false  # Restoration failed
        }
    } catch {
        LogMessage "An error occurred while restoring ${User}'s browser settings from: $BackupDir to: $ProfileDir. Error: $_"
        return $false  # Restoration failed

    }
}

# Main script execution
try {
    # Get list of user directories from backup location
    $backupBaseDir = "\\${siteNumber}-pcName\Redacted\Browser Settings Backups\${hostname}"

    # Check for backup directory accessibility
    if (!(Test-Path -Path $backupBaseDir)) {
        throw "Backup base directory does not exist or is inaccessible: $backupBaseDir"
    }

    # Check for backup directory corruption or inaccessibility
    if (CheckBackupDirCorruption -BackupDir $backupBaseDir) {
        throw "Backup base directory is corrupted or inaccessible: $backupBaseDir"
    }

    $userDirs = Get-ChildItem -Path $backupBaseDir -Directory
    $userList = $userDirs.Name

    # Check that $userList is not empty
    if ($userList.Count -eq 0) {
        throw "No user directories found in backup location: $backupBaseDir"
    }

    # Initialize lists to track successful and failed restorations for Chrome and Edge separately
    $successfulChromeRestoration = @()
    $failedChromeRestoration = @()
    $successfulEdgeRestoration = @()
    $failedEdgeRestoration = @()
    $successfulChromePermissionReset = @()
    $failedChromePermissionReset = @()
    $successfulEdgePermissionReset = @()
    $failedEdgePermissionReset = @()
    $successfulTotalRestore = @()

    # Iterate through each user to restore browser settings for Chrome and Edge separately
    foreach ($user in $userList) {
        $userProfileDir = "C:\Users\$user"
        $userBackupDir = Join-Path -Path $backupBaseDir -ChildPath $user

        # Check if the user profile directory exists
        if (Test-Path -Path $userProfileDir) {
            # Define profile directories for Chrome and Edge
            $chromeProfileDir = Join-Path -Path $userProfileDir -ChildPath "AppData\Local\Google\Chrome\User Data\Default"
            $edgeProfileDir = Join-Path -Path $userProfileDir -ChildPath "AppData\Local\Microsoft\Edge\User Data\Default"

            # Define backup directories for Chrome and Edge
            $chromeBackupDir = Join-Path -Path $userBackupDir -ChildPath "Chrome"
            $edgeBackupDir = Join-Path -Path $userBackupDir -ChildPath "Edge"

            # Check for sufficient disk space before proceeding
            $requiredSpace = (Get-ChildItem -Path $chromeBackupDir -Recurse | Measure-Object -Property Length -Sum).Sum + 
                             (Get-ChildItem -Path $edgeBackupDir -Recurse | Measure-Object -Property Length -Sum).Sum
            $freeSpace = [System.IO.DriveInfo]::GetDriveInfo("C:").AvailableFreeSpace

            if ($freeSpace -lt $requiredSpace) {
                LogMessage "Insufficient disk space for user ${user}. Required: $requiredSpace, Available: $freeSpace"
                $failedChromeRestoration += $user
                $failedEdgeRestoration += $user
                continue
            }

            # Restore Chrome settings and update lists accordingly
            if (RestoreBrowserSettings -User $user -ProfileDir $chromeProfileDir -BackupDir $chromeBackupDir -FilesToRestore @("Bookmarks", "Preferences", "Login Data", "History", "Cookies", "Web Data") -FoldersToRestore @("Extensions", "Local Storage", "Session Storage", "Sync Data")) {
                $successfulChromeRestoration += $user
            } else {
                $failedChromeRestoration += $user
            }

            # Restore Edge settings and update lists accordingly
            if (RestoreBrowserSettings -User $user -ProfileDir $edgeProfileDir -BackupDir $edgeBackupDir -FilesToRestore @("Bookmarks", "Preferences", "Login Data", "History", "Cookies", "Web Data") -FoldersToRestore @("Extensions", "Local Storage", "Session Storage", "Sync Data")) {
                $successfulEdgeRestoration += $user
            } else {
                $failedEdgeRestoration += $user
            }

            # Reset permissions for Chrome and update lists accordingly
            if (ResetPermissions -User $user -Path $chromeProfileDir) {
                $successfulChromePermissionReset += $user
            } else {
                $failedChromePermissionReset += $user
            }
            # Reset permissions for Edge and update lists accordingly
            if (ResetPermissions -User $user -Path $edgeProfileDir) {
                $successfulEdgePermissionReset += $user
            } else {
                $failedEdgePermissionReset += $user
            }
        } else {
            LogMessage "User profile directory does not exist: ${userProfileDir}. Skipping user: ${user}."
            $failedChromeRestoration += $user
            $failedEdgeRestoration += $user
        }
    }

    # Generate list of users that passed every step
    # Combine all success and failure arrays
    $allSuccessful = $successfulChromeRestoration + $successfulEdgeRestoration + $successfulChromePermissionReset + $successfulEdgePermissionReset
    $allFailed = $failedChromeRestoration + $failedEdgeRestoration + $failedChromePermissionReset + $failedEdgePermissionReset

    # Filter out users who have failed any of the steps
    $successfulTotalRestore = $allSuccessful | Where-Object {
        $user = $_
        $allFailed -notcontains $user
    }

    # Remove duplicates
    $successfulTotalRestore = $successfulTotalRestore | Select-Object -Unique

    #Log and display results
    if ($successfulTotalRestore.Count -gt 0) {
        $successMessage = "Successfully restored Chrome and Edge settings for: $($successfulTotalRestore -join ', ')."
        LogMessage $successMessage
        Write-Host $successMessage
    }

    if ($failedChromeRestoration.Count -gt 0) {
        $failedRestoreMessage = "Failed to restore Chrome settings for: $($failedChromeRestoration -join ', ')."
        LogMessage $failedRestoreMessage
        Write-Host $failedRestoreMessage
    }

    if ($failedEdgeRestoration.Count -gt 0) {
        $failedRestoreMessage = "Failed to restore Edge settings for: $($failedEdgeRestoration -join ', ')."
        LogMessage $failedRestoreMessage
        Write-Host $failedRestoreMessage
    }

    if ($failedChromePermissionReset.Count -gt 0) {
        $failedResetMessage = "Failed to reset Chrome file permissions for: $($failedChromePermissionReset -join ', ')."
        LogMessage $failedResetMessage
        Write-Host $failedResetMessage
    }

    if ($failedEdgePermissionReset.Count -gt 0) {
        $failedResetMessage = "Failed to reset Edge file permissions for: $($failedEdgePermissionReset -join ', ')."
        LogMessage $failedResetMessage
        Write-Host $failedResetMessage
    }

} catch {
    LogMessage "An error occurred: $_"
    Write-Host "An error occurred: $_"
}

Start-Sleep -Seconds 10
Exit