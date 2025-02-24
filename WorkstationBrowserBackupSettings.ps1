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
        } catch {
            LogMessage "Unexpected error occurred: $_. Exiting operation." "ERROR"
            return $false
        }
    }
    LogMessage "Operation failed after $retryCount retries due to network interruptions."
    return $false
}

# Function to check for directory corruption
function CheckDirCorruption {
    param (
        [string]$Dir
    )
    if (!(Test-Path -Path $Dir -ErrorAction SilentlyContinue)) {
        LogMessage "$'{Dir}' doesn't exist, assumed corrupted" "ERROR"
        return $true
    }
    # Check for actual corruption or inaccessibility
    try {
        Get-ChildItem $Dir -ErrorAction Stop | Out-Null
        LogMessage "'${Dir}' is intact and accessible"
        return $false
    } catch {
        LogMessage "'${Dir}' is corrupted or inaccessible" "ERROR"
        return $true
    }
}

# Function to create backup directory
function CreateBackupDirectory {
    param (
        [string]$Path
    )
    try {
        if (!(Test-Path -Path $Path)) {
            New-Item -ItemType Directory -Path $Path | Out-Null
            if (Test-Path -Path $Path) {
                LogMessage "Backup directory created: $Path"
            } else {
                LogMessage "Failed to create backup directory: $Path" "ERROR"
            }
        } else {
            LogMessage "Backup directory already exists: $Path"
        }
    } catch {
        LogMessage "Failed to create or verify backup directory: $Path. Error: $_" "ERROR"
    }
}

# Initialize lists to track successful and failed backups for Chrome and Edge separately
$successfulChromeBackup = @()
$failedChromeBackup = @()
$successfulEdgeBackup = @()
$failedEdgeBackup = @()

# Test if existing document repository is accessible
$ExistingDocumentRepository = "\\${siteNumber}-pcName\Redacted\"
if (!(Test-Path -Path $ExistingDocumentRepository)) {
    LogMessage "$ExistingDocumentRepository is inaccessible, exiting script..."
    Exit
}

# Create the backup directory in existing document repository
$backupDir = "\\${siteNumber}-pcName\Redacted\Browser Settings Backups\${hostname}"
if (-not (HandleNetworkInterruptions { CreateBackupDirectory -Path $backupDir })) {
    LogMessage "Failed to create backup directory $backupDir after multiple attempts. Exiting script." "ERROR"
    Exit
}

# Get list of user directories
$usersDir = Get-ChildItem -Path "C:\Users" -Directory

# Iterate through users directory and add names to the user list
$userList = $usersDir.Name | Where-Object { $_ -notin @("Administrator", "Public") }

# Function to back up browser settings
function BackupBrowserSettings {
    param (
        [string]$ProfileDir,
        [string]$BackupDir,
        [string[]]$FilesToBackup,
        [string[]]$FoldersToBackup
    )
    $backupSuccess = $true
    if (Test-Path -Path $ProfileDir) {
        if (-not (HandleNetworkInterruptions { CreateBackupDirectory -Path $BackupDir })) {
            LogMessage "Failed to create backup directory $BackupDir after multiple attempts. Skipping backup for $ProfileDir." "ERROR"
            return
        }
        foreach ($file in $FilesToBackup) {
            $sourcePath = Join-Path -Path $ProfileDir -ChildPath $file

            # Check if file exists in source directory before copying
            if (Test-Path -Path $sourcePath) {
                $copyOperation = {
                    Copy-Item -Path $sourcePath -Destination $BackupDir -Recurse -Force
                }
                if (-not (HandleNetworkInterruptions $copyOperation)) {
                    LogMessage "Failed to copy file $sourcePath to $BackupDir after multiple attempts." "ERROR"
                    $backupSuccess = $false
                } else {
                    LogMessage "Successfully backed up file $sourcePath to $BackupDir."
                }
            }
        }
        foreach ($folder in $FoldersToBackup) {
            $sourcePath = Join-Path -Path $ProfileDir -ChildPath $folder

            # Check if folder exists in source directory before copying
            if (Test-Path -Path $sourcePath) {
                $destinationPath = Join-Path -Path $BackupDir -ChildPath $folder
                $copyOperation = {
                    Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse -Force
                }
                if (-not (HandleNetworkInterruptions $copyOperation)) {
                    LogMessage "Failed to copy folder $sourcePath to $destinationPath after multiple attempts." "ERROR"
                    $backupSuccess = $false
                } else {
                    LogMessage "Successfully backed up folder $sourcePath to $destinationPath."
                }
            }
        }
    } else {
        LogMessage "Profile directory $ProfileDir does not exist. Skipping backup." "ERROR"
        $backupSuccess = $false
    }
    return $backupSuccess
}

# Iterate through each user to back up browser settings
foreach ($user in $userList) {
    $userBackupDir = Join-Path -Path $backupDir -ChildPath $user
    if (-not (HandleNetworkInterruptions { CreateBackupDirectory -Path $userBackupDir })) {
        LogMessage "Failed to create user backup directory $userBackupDir after multiple attempts. Skipping user $user." "ERROR"
        continue
    }

    # Define backup directories for Chrome and Edge
    $chromeBackupDir = Join-Path -Path $userBackupDir -ChildPath "Chrome"
    $edgeBackupDir = Join-Path -Path $userBackupDir -ChildPath "Edge"

    # Define profile directories for Chrome and Edge
    $chromeProfileDir = "C:\Users\${user}\AppData\Local\Google\Chrome\User Data\Default"
    $edgeProfileDir = "C:\Users\${user}\AppData\Local\Microsoft\Edge\User Data\Default"

    # Check for directory corruption before proceeding with backups
    if (-not (CheckDirCorruption -Dir $chromeProfileDir)) {
        # Back up Chrome settings
        if (BackupBrowserSettings -User $user -ProfileDir $chromeProfileDir -BackupDir $chromeBackupDir -FilesToBackup @("Bookmarks", "Preferences", "Login Data", "History", "Cookies", "Web Data") -FoldersToBackup @("Extensions", "Local Storage", "Session Storage", "Sync Data")) {
            $successfulChromeBackup += $user
        } else {
            $failedChromeBackup += $user
        }
    } else {
        LogMessage "Chrome profile directory $chromeProfileDir is corrupted or inaccessible. Skipping backup." "ERROR"
        $failedChromeBackup += $user
    }
    
    if (-not (CheckDirCorruption -Dir $edgeProfileDir)) {
        # Back up Edge settings
        if (BackupBrowserSettings -User $user -ProfileDir $edgeProfileDir -BackupDir $edgeBackupDir -FilesToBackup @("Bookmarks", "Preferences", "Login Data", "History", "Cookies", "Web Data") -FoldersToBackup @("Extensions", "Local Storage", "Session Storage", "Sync Data")) {
            $successfulEdgeBackup += $user
        } else {
            $failedEdgeBackup += $user
        }
    } else {
        LogMessage "Edge profile directory $edgeProfileDir is corrupted or inaccessible. Skipping backup." "ERROR"
        $failedEdgeBackup += $user
    }
}

# Combine all success and failure arrays
$allSuccessful = $successfulChromeBackup + $successfulEdgeBackup
$allFailed = $failedChromeBackup + $failedEdgeBackup

# Filter out users who have failed any of the steps
$successfulTotalBackup = $allSuccessful | Where-Object {
    $user = $_
    $allFailed -notcontains $user
}

# Remove duplicates
$successfulTotalBackup = $successfulTotalBackup | Select-Object -Unique

# Log and display results
if ($successfulTotalBackup.Count -gt 0) {
    $successMessage = "Successfully backed up Chrome and Edge settings for: $($successfulTotalBackup -join ', ')."
    LogMessage $successMessage
    Write-Host $successMessage
}

if ($failedChromeBackup.Count -gt 0) {
    $failedMessage = "Failed to back up Chrome settings for: $($failedChromeBackup -join ', ')."
    LogMessage $failedMessage
    Write-Host $failedMessage
}

if ($failedEdgeBackup.Count -gt 0) {
    $failedMessage = "Failed to back up Edge settings for: $($failedEdgeBackup -join ', ')."
    LogMessage $failedMessage
    Write-Host $failedMessage
}

LogMessage "Browser settings have been backed up!"
Start-Sleep -Seconds 10