
<#
Requires: Hyper-V module
Run as: Administrator
Purpose: Sequentially stop -> export -> start multiple VMs (default AD01/AD02),
         then wait 30 minutes and ping target before proceeding to next VM.
Auditing: Transcript -> HyperV-Export-Transcript.log
          Summary lines -> HyperV-Export-Audit.log   (ISO 9001 traceability)
#>

param(
    [string[]]$VMNames = @("AD01","AD02"),
    [hashtable]$PingTargets = @{ "AD01" = "AD01"; "AD02" = "AD02" },

    [string]$RootExportPath = "D:\HyperV_Monthly_Backups",
    [int]$ShutdownTimeoutSec = 180,
    [int]$MinFreeSpaceGB = 20,

    [int]$PostStartWaitSeconds = 1800,   # 30 minutes
    [int]$PingCount = 4,
    [int]$PingRetries = 3,
    [int]$PingRetryIntervalSec = 20
)

# --- Preconditions ---
if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
    Write-Error "Hyper-V module not found. Please enable/install Hyper-V management tools."
    exit 1
}
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Error "Run this script as Administrator."
    exit 1
}

# --- Paths ---
if (-not (Test-Path $RootExportPath)) {
    New-Item -Path $RootExportPath -ItemType Directory | Out-Null
}
$CurrentDate      = Get-Date -Format "yyyy-MM-dd"
$TargetExportPath = Join-Path -Path $RootExportPath -ChildPath $CurrentDate
if (-not (Test-Path $TargetExportPath)) {
    New-Item -Path $TargetExportPath -ItemType Directory -Force | Out-Null
}

# Separate files to avoid locking
$TranscriptFile = Join-Path -Path $RootExportPath -ChildPath "HyperV-Export-Transcript.log"
$AuditFile      = Join-Path -Path $RootExportPath -ChildPath "HyperV-Export-Audit.log"

# --- Transcript ---
try {
    Start-Transcript -Path $TranscriptFile -Append -ErrorAction Stop
    Write-Host ("Transcript started. Appending logs to: {0}" -f $TranscriptFile)
} catch {
    Write-Warning ("Transcript could not start: {0}" -f $_.Exception.Message)
}

function Get-EstimatedVmSizeGB([string]$VMName) {
    try {
        $paths = Get-VMHardDiskDrive -VMName $VMName -ErrorAction Stop | Select-Object -ExpandProperty Path
        $totalBytes = 0
        foreach ($p in $paths) { if (Test-Path $p) { $totalBytes += (Get-Item $p).Length } }
        return [math]::Round($totalBytes/1GB, 2)
    } catch {
        Write-Warning ("Could not estimate VHD size for '{0}': {1}" -f $VMName, $_.Exception.Message)
        return 0
    }
}

function Test-FreeSpace([string]$Path, [int]$MinGB, [double]$EstimatedGB) {
    $drive       = Get-Item $Path
    $driveLetter = $drive.PSDrive.Name
    $vol         = Get-PSDrive -Name $driveLetter
    $freeGB      = [math]::Round($vol.Free/1GB, 2)
    Write-Host ("Free space on drive {0}: {1} GB (Estimated VM size: {2} GB)" -f $driveLetter, $freeGB, $EstimatedGB)
    $needGB = $MinGB + $EstimatedGB
    if ($freeGB -lt $needGB) {
        throw ("Insufficient free space (< {0} GB required). Current: {1} GB." -f $needGB, $freeGB)
    }
    return $freeGB
}

function Stop-VMGracefully([string]$VMName, [int]$TimeoutSec) {
    $vm = Get-VM -Name $VMName -ErrorAction Stop
    if ($vm.State -eq 'Running') {
        Write-Host ("Requesting graceful shutdown: '{0}'..." -f $VMName)
        Stop-VM -Name $VMName -ErrorAction Stop
    } else {
        Write-Host ("VM '{0}' not in 'Running' state (state: {1}). Proceeding." -f $VMName, $vm.State)
    }

    # Wait until Off, fallback to TurnOff if needed
    $elapsed = 0
    while ((Get-VM -Name $VMName).State -ne 'Off' -and $elapsed -lt $TimeoutSec) {
        Start-Sleep -Seconds 3
        $elapsed += 3
    }
    if ((Get-VM -Name $VMName).State -ne 'Off') {
        Write-Warning ("Graceful shutdown timed out. Forcing turn-off for '{0}'..." -f $VMName)
        Stop-VM -Name $VMName -TurnOff -ErrorAction Stop
        while ((Get-VM -Name $VMName).State -ne 'Off') { Start-Sleep -Seconds 2 }
    }
}

function Start-VMAndWait([string]$VMName) {
    Write-Host ("Starting VM '{0}'..." -f $VMName)
    Start-VM -Name $VMName -ErrorAction Stop | Out-Null
    Write-Host ("VM '{0}' started successfully." -f $VMName)
}

function WaitAndPing([string]$VMName, [string]$PingTarget, [int]$WaitSec, [int]$Count, [int]$Retries, [int]$RetryIntervalSec) {
    Write-Host ("Waiting {0} seconds before health check for '{1}' ({2})..." -f $WaitSec, $VMName, $PingTarget)
    Start-Sleep -Seconds $WaitSec

    for ($i = 1; $i -le $Retries; $i++) {
        Write-Host ("Ping attempt {0}/{1}: {2}" -f $i, $Retries, $PingTarget)
        try {
            $ok = Test-Connection -ComputerName $PingTarget -Count $Count -Quiet -ErrorAction Stop
        } catch {
            $ok = $false
            Write-Warning ("Ping failed (exception): {0}" -f $_.Exception.Message)
        }
        if ($ok) {
            Write-Host ("Ping successful for '{0}'. Proceeding to next VM." -f $PingTarget)
            return $true
        } else {
            if ($i -lt $Retries) {
                Write-Host ("Ping not successful. Will retry in {0} seconds..." -f $RetryIntervalSec)
                Start-Sleep -Seconds $RetryIntervalSec
            }
        }
    }
    Write-Warning ("Ping unsuccessful after {0} attempts for '{1}'." -f $Retries, $PingTarget)
    return $false
}

foreach ($VMName in $VMNames) {
    # Validate VM
    $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
    if (-not $vm) {
        Write-Error ("VM '{0}' not found on this host. Skipping." -f $VMName)
        continue
    }

    # Prepare per-VM export path; avoid overwrite if already exists
    $VmExportPath = Join-Path -Path $TargetExportPath -ChildPath $VMName
    if (Test-Path $VmExportPath) {
        $stamp = Get-Date -Format "HHmmss"
        $VmExportPath = Join-Path -Path $TargetExportPath -ChildPath ("{0}-{1}" -f $VMName, $stamp)
        Write-Warning ("Existing VM export folder detected; using new path: '{0}'" -f $VmExportPath)
    }
    New-Item -Path $VmExportPath -ItemType Directory -Force | Out-Null

    # Space check (with estimate)
    $estimateGB = Get-EstimatedVmSizeGB -VMName $VMName
    try {
        $null = Test-FreeSpace -Path $RootExportPath -MinGB $MinFreeSpaceGB -EstimatedGB $estimateGB
    } catch {
        Write-Error $_.Exception.Message
        $RunEndTime = Get-Date
        $AuditLine  = ("[{0}] Host={1} User={2} VM={3} Result={4} ExportPath={5}" -f ($RunEndTime.ToString("yyyy-MM-dd HH:mm:ss")), $env:COMPUTERNAME, $env:USERNAME, $VMName, "Failure: Insufficient space", $VmExportPath)
        Add-Content -Path $AuditFile -Value $AuditLine
        continue
    }

    $Outcome     = "Unknown"
    $RunStart    = Get-Date
    $HostName    = $env:COMPUTERNAME
    $RunUser     = $env:USERNAME

    try {
        Stop-VMGracefully -VMName $VMName -TimeoutSec $ShutdownTimeoutSec

        Write-Host ("Exporting VM '{0}' to '{1}'..." -f $VMName, $VmExportPath)
        Export-VM -Name $VMName -Path $VmExportPath -ErrorAction Stop
        Write-Host ("Export completed for '{0}'." -f $VMName)

        Start-VMAndWait -VMName $VMName
        $Outcome = "Success"
    } catch {
        Write-Error ("Export workflow failed for '{0}'. Error: {1}" -f $VMName, $_.Exception.Message)
        $Outcome = ("Failure: {0}" -f $_.Exception.Message)

        try {
            if ((Get-VM -Name $VMName).State -eq 'Off') {
                Write-Warning ("Attempting to start VM '{0}' after failure..." -f $VMName)
                Start-VMAndWait -VMName $VMName
            }
        } catch {}
    } finally {
        $RunEnd = Get-Date
        $AuditLine = ("[{0}] Host={1} User={2} VM={3} Result={4} ExportPath={5} Start={6} End={7}" -f ($RunEnd.ToString("yyyy-MM-dd HH:mm:ss")), $HostName, $RunUser, $VMName, $Outcome, $VmExportPath, $RunStart.ToString("HH:mm:ss"), $RunEnd.ToString("HH:mm:ss"))
        Add-Content -Path $AuditFile -Value $AuditLine
    }

    # Inter-VM wait & health check before proceeding
    $currentIndex = [array]::IndexOf($VMNames, $VMName)
    if ($currentIndex -lt ($VMNames.Length - 1)) {
        $nextVM = $VMNames[$currentIndex + 1]

        $pingTarget = $VMName
        if ($PingTargets -and $PingTargets.ContainsKey($VMName)) { $pingTarget = $PingTargets[$VMName] }

        $ok = WaitAndPing -VMName $VMName -PingTarget $pingTarget -WaitSec $PostStartWaitSeconds -Count $PingCount -Retries $PingRetries -RetryIntervalSec $PingRetryIntervalSec
        if (-not $ok) {
            $msg = ("Health check did not pass for '{0}' ({1}). Will NOT proceed to next VM '{2}'." -f $VMName, $pingTarget, $nextVM)
            Write-Warning $msg
            $AuditLineHC = ("[{0}] Host={1} User={2} VM={3} Result={4}" -f ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")), $HostName, $RunUser, $VMName, $msg)
            Add-Content -Path $AuditFile -Value $AuditLineHC
            break
        } else {
            Write-Host ("Health check OK for '{0}'. Proceeding to next VM '{1}'..." -f $VMName, $nextVM)
        }
    }
}

try { Stop-Transcript | Out-Null } catch {}
Write-Host ("All steps finished. Root destination: {0}" -f $TargetExportPath)
``
