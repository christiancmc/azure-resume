<#
.NAME
    DellDrivers_Installation

.VERSION
    1.3.2

.SYNOPSIS
    Script to install Dell Command Update and run driver updates on remote computers using PsExec (Local System).

.DESCRIPTION
    - Maintains a list of computers in C:\Scripts\computers_list.txt.
    - Allows adding/removing computers.
    - Checks whether computers are online (via ping) and verifies that the last 3 lines of
      C:\SWDEPOT\LOGS\OSDResults.log contain:
          • "NO Device profile provisioning found for this device."
          • "Removing schedule task"
          • "Rebooting System"
    - Once those conditions are met, the script waits 30 seconds before proceeding.
    - It then copies Dell-Command-Update.exe to the remote computer (in C:\SWDEPOT).
    - Installs Dell Command Update silently using PsExec in Local System mode.
    - Waits 5 seconds after successful installation before proceeding.
    - Executes DCU-CLI with /applyUpdates -reboot=enable -silent three times.
         If any iteration returns a nonzero exit code, it logs the error and immediately
         continues to the next iteration.
    - Finally, it issues a remote reboot command.
    - Offline or “log not ready” messages are limited to 2 per computer per pass,
      and ready computers are given priority.
     
.NOTES
    - Requires PsExec.exe in C:\Tools.
    - The Dell Command Update installer must be at C:\Drivers\Dell-Command-Update.exe.
    - The user running this script must be an Administrator on the remote computers.
    - Uses the Local System account on the remote computer (via '-s'), so no credentials are prompted.
    - If BitLocker is enabled on remote systems, a BIOS update might trigger a recovery key prompt unless BitLocker is suspended.
#>

# ---------------------------
# --- GLOBAL VARIABLES ---
# ---------------------------
$Global:Computers = @()
$Global:ComputerListFile = "C:\Scripts\computers_list.txt"

# Log file requirements
$Global:LogPathRelative = "C$\SWDEPOT\LOGS\OSDResults.log"
$Global:LineContains1   = "NO Device profile provisioning found for this device."
$Global:LineContains2   = "Removing schedule task"
$Global:LineContains3   = "Rebooting System"

# Paths
$Global:PsExecPath = "C:\Tools\PsExec.exe"
$Global:DellCmdUpdateLocal = "C:\Drivers\Dell-Command-Update.exe"

# ---------------------------
# --- FUNCTIONS ---
# ---------------------------
function LoadComputerList {
    $Global:Computers = @()
    if (Test-Path $Global:ComputerListFile) {
        Write-Host "DEBUG: Loading list from $Global:ComputerListFile"
        $lines = Get-Content -Path $Global:ComputerListFile -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }
        if ($lines) { $Global:Computers = $lines }
        Write-Host "DEBUG: $($Global:Computers.Count) computer(s) loaded: $($Global:Computers -join ' | ')"
    }
    else {
        Write-Host "DEBUG: No computer list file found at $Global:ComputerListFile"
    }
}

function SaveComputerList {
    $scriptsFolder = Split-Path $Global:ComputerListFile
    if (-not (Test-Path $scriptsFolder)) { New-Item -ItemType Directory -Path $scriptsFolder -Force | Out-Null }
    Write-Host "DEBUG: Saving computers (count=$($Global:Computers.Count)): $($Global:Computers -join ' | ')"
    $Global:Computers | Set-Content -Path $Global:ComputerListFile -Encoding UTF8
}

function ShowMenu {
    Clear-Host
    Write-Host "==========================================="
    Write-Host " Dell Drivers Installation (v1.3.2)"
    Write-Host "==========================================="
    Write-Host "1. Add computers"
    Write-Host "2. Remove computers"
    Write-Host "3. Show computer list and online/offline status"
    Write-Host "4. Start driver installation process"
    Write-Host "5. Exit"
    Write-Host "==========================================="
}

function AddComputers {
    $newEntries = Read-Host "Enter the computer name(s) or IP(s) (separated by spaces or commas)"
    Write-Host "DEBUG: You typed: [$newEntries]"
    $parsedList = $newEntries -split "[,\s]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    Write-Host "DEBUG: Split items = $($parsedList -join ' | ')"
    foreach ($pc in $parsedList) {
        if ($null -eq $Global:Computers) { $Global:Computers = @() }
        if (-not $Global:Computers.Contains($pc)) {
            $Global:Computers += $pc
            Write-Host "Computer '$pc' added to the list."
        } else {
            Write-Host "Computer '$pc' is already in the list."
        }
    }
    SaveComputerList
    Pause
}

function RemoveComputers {
    if ($Global:Computers.Count -eq 0) {
        Write-Host "There are no computers in the list to remove."
        Pause
        return
    }
    Write-Host "`nCurrent computers in the list:"
    foreach ($pc in $Global:Computers) { Write-Host "- $pc" }
    $toRemove = Read-Host "`nEnter the computer name(s) or IP(s) to remove, separated by commas/spaces"
    $parsedRemoveList = $toRemove -split "[,\s]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    foreach ($pc in $parsedRemoveList) {
        if ($Global:Computers.Contains($pc)) {
            $Global:Computers = $Global:Computers | Where-Object { $_ -ne $pc }
            Write-Host "Computer '$pc' removed from the list."
        } else {
            Write-Host "Computer '$pc' is not in the list."
        }
    }
    SaveComputerList
    Pause
}

function ShowComputerStatus {
    if ($Global:Computers.Count -eq 0) {
        Write-Host "No computers in the list."
    }
    else {
        foreach ($pc in $Global:Computers) {
            $online = Test-Connection -ComputerName $pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            if ($online) { Write-Host "$pc - ONLINE" -ForegroundColor Green }
            else { Write-Host "$pc - OFFLINE" -ForegroundColor Red }
        }
    }
    Pause
}

#---------------------------------------------------------------------
# Function: Test-ComputerReadiness
# Checks if the remote PC is online and if its OSDResults.log's last 3 lines contain the required text.
# Returns $true if ready, $false otherwise.
#---------------------------------------------------------------------
function Test-ComputerReadiness($ComputerName) {
    $online = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $online) { return $false }
    $remoteLogPath = "\\$ComputerName\$($Global:LogPathRelative)"
    if (Test-Path $remoteLogPath) {
        $last3 = Get-Content $remoteLogPath -Tail 3
        if ($last3.Count -ge 3) {
            $cond1 = ($last3[0] -like "*$($Global:LineContains1)*")
            $cond2 = ($last3[1] -like "*$($Global:LineContains2)*")
            $cond3 = ($last3[2] -like "*$($Global:LineContains3)*")
            if ($cond1 -and $cond2 -and $cond3) { return $true }
        }
    }
    return $false
}

#---------------------------------------------------------------------
# Function: InstallDellDrivers
# Performs:
#   1) Waits for the remote log condition, then waits an additional 30 seconds.
#   2) Copies Dell-Command-Update.exe to the remote machine.
#   3) Installs Dell Command Update silently using PsExec.
#      After installation, waits 5 seconds before proceeding.
#   4) Executes DCU-CLI with /applyUpdates -reboot=enable -silent three times.
#      If an iteration returns a nonzero exit code, it logs the error and continues.
#   5) Issues a remote reboot command.
#---------------------------------------------------------------------
function InstallDellDrivers($ComputerName) {
    Write-Host "`n==========================================="
    Write-Host " Starting process on $ComputerName..."
    Write-Host "==========================================="

    # --- Step 1: Wait for OSDResults.log condition ---
    $remoteLogPath = "\\$ComputerName\$($Global:LogPathRelative)"
    while ($true) {
        if (Test-Path $remoteLogPath) {
            $lastLines = Get-Content $remoteLogPath -Tail 3
            if ($lastLines.Count -ge 3) {
                $condition1 = ($lastLines[0] -like "*$($Global:LineContains1)*")
                $condition2 = ($lastLines[1] -like "*$($Global:LineContains2)*")
                $condition3 = ($lastLines[2] -like "*$($Global:LineContains3)*")
                if ($condition1 -and $condition2 -and $condition3) {
                    Write-Host "OSDResults.log found and last 3 lines match the required text."
                    break
                }
            }
        }
        Write-Host "Remote log not ready or doesn't match the required text yet. Retrying in 5s..."
        Start-Sleep -Seconds 5
    }
    Write-Host "Conditions met. Waiting 30 seconds before proceeding..." -ForegroundColor Yellow
    Start-Sleep -Seconds 30

    # --- Step 2: Copy Dell-Command-Update.exe ---
    $remoteSWDepot = "\\$ComputerName\C$\SWDEPOT"
    Write-Host "Creating folder at $remoteSWDepot (if needed) and copying the installer..."
    New-Item -ItemType Directory -Path $remoteSWDepot -Force -ErrorAction SilentlyContinue | Out-Null
    $localFilePath  = $Global:DellCmdUpdateLocal
    $remoteFilePath = Join-Path $remoteSWDepot "Dell-Command-Update.exe"
    try {
        Copy-Item -Path $localFilePath -Destination $remoteFilePath -Force -ErrorAction Stop
        Write-Host "File successfully copied to $remoteSWDepot."
    }
    catch {
        Write-Host "Error copying the file: $($_)" -ForegroundColor Red
        return
    }

    # --- Step 3: Install Dell Command Update silently ---
    Write-Host "Installing Dell Command Update silently on $ComputerName (Local System via PsExec)..."
    Write-Host "This may take a few minutes. Please do not close this window..." -ForegroundColor Green
    $startTime = Get-Date
    $remoteCmd = "C:\SWDEPOT\Dell-Command-Update.exe /s"
    $psexecArgs = @(
        "\\$ComputerName"
        "-s"
        "-h"
        "-accepteula"
        $remoteCmd
    )
    try {
        Start-Process -FilePath $Global:PsExecPath -ArgumentList $psexecArgs -Wait -NoNewWindow -ErrorAction Stop
        $endTime = Get-Date
        $duration = $endTime - $startTime
        Write-Host "Dell Command Update installed successfully on $ComputerName."
        Write-Host "Duration: $($duration.ToString())"
    }
    catch {
        Write-Host "Error installing Dell Command Update on $($ComputerName): $($_)" -ForegroundColor Red
        return
    }

    # --- New Delay: Wait 5 seconds before proceeding ---
    Write-Host "Waiting 5 seconds before executing DCU-CLI..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    # --- Step 4: Run DCU-CLI three times ---
    Write-Host "Running DCU-CLI to update drivers (including BIOS) on $ComputerName..."
    Write-Host "This step can also take several minutes (or more). Please do not close this window..." -ForegroundColor Green
    $dcuCliRemote = '"C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" /applyUpdates -reboot=enable -silent'
    $psexecArgsDCU = @(
        "\\$ComputerName"
        "-s"
        "-h"
        "-accepteula"
        $dcuCliRemote
    )
    for ($i = 1; $i -le 3; $i++) {
        Write-Host "Running DCU-CLI iteration $i on $ComputerName..." -ForegroundColor Green
        $iterStart = Get-Date
        try {
            $proc = Start-Process -FilePath $Global:PsExecPath -ArgumentList $psexecArgsDCU -NoNewWindow -PassThru -Wait -ErrorAction Stop
            if ($proc.ExitCode -ne 0) {
                Write-Host "Iteration $i on $ComputerName returned exit code $($proc.ExitCode). Continuing to next iteration." -ForegroundColor Red
                continue
            }
            $iterEnd = Get-Date
            $iterDuration = $iterEnd - $iterStart
            Write-Host "Iteration $i completed on $ComputerName. Duration: $($iterDuration.ToString())"
        }
        catch {
            Write-Host "Error in DCU-CLI iteration $i on $($ComputerName): $($_)" -ForegroundColor Red
            continue
        }
    }

    # --- Step 5: Issue remote reboot command ---
    Write-Host "Initiating remote reboot for $ComputerName..." -ForegroundColor Green
    $rebootCmd = "shutdown /r /t 0"
    $rebootArgs = @(
        "\\$ComputerName",
        "-s",
        "-h",
        "-accepteula",
        $rebootCmd
    )
    try {
        Start-Process -FilePath $Global:PsExecPath -ArgumentList $rebootArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "Remote reboot command issued for $ComputerName."
    }
    catch {
        Write-Host "Error issuing remote reboot command on $($ComputerName): $($_)" -ForegroundColor Red
    }

    Write-Host "==========================================="
    Write-Host " Process completed for $ComputerName."
    Write-Host "==========================================="
}

#---------------------------------------------------------------------
# Main installation loop with priority:
# - We keep a list of remaining computers to process.
# - In each pass, if a computer is ready (online + log condition), we install it immediately.
# - For those not ready, we print up to 2 "offline" or "log not ready" messages.
# - We loop until all computers are processed.
#---------------------------------------------------------------------
function StartDriverInstallation {
    if ($Global:Computers.Count -eq 0) {
        Write-Host "No computers in the list. Please add at least one computer."
        Pause
        return
    }
    $remaining = New-Object System.Collections.ArrayList
    [void] $remaining.AddRange($Global:Computers)
    $offlineCount = @{}
    $logCount     = @{}
    foreach ($pc in $remaining) {
        $offlineCount[$pc] = 0
        $logCount[$pc]     = 0
    }
    while ($remaining.Count -gt 0) {
        $processedAny = $false
        $currentList = $remaining.Clone()
        foreach ($pc in $currentList) {
            $ready = Test-ComputerReadiness $pc
            if ($ready) {
                InstallDellDrivers $pc
                [void] $remaining.Remove($pc)
                $processedAny = $true
            }
            else {
                $isOnline = Test-Connection -ComputerName $pc -Count 1 -Quiet -ErrorAction SilentlyContinue
                if (-not $isOnline) {
                    if ($offlineCount[$pc] -lt 2) {
                        Write-Host "Computer $pc is offline. Skipping for now..."
                        $offlineCount[$pc]++
                    }
                }
                else {
                    if ($logCount[$pc] -lt 2) {
                        Write-Host "Computer $pc is online but OSDResults.log is not ready. Skipping for now..."
                        $logCount[$pc]++
                    }
                }
            }
        }
        if (-not $processedAny) { Start-Sleep -Seconds 5 } else { continue }
    }
    Write-Host "`nAll computers have been processed (or removed)."
    Pause
}

# ---------------------------
# Script Entry Point
# ---------------------------
LoadComputerList
do {
    ShowMenu
    $option = Read-Host "Select an option [1-5]"
    switch ($option) {
        1 { AddComputers }
        2 { RemoveComputers }
        3 { ShowComputerStatus }
        4 { StartDriverInstallation }
        5 { Write-Host "Exiting script..." }
        default { Write-Host "Invalid option. Please try again."; Pause }
    }
} while ($option -ne 5)
