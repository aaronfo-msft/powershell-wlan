using namespace System.Diagnostics;

class ProgressTracker {
    [string]$Activity
    [string]$CurrentOperation = $null
    [string]$Status = $null
    [DateTime]$StartTime
    [DateTime]$LastUpdateTime
    [int]$Total

    ProgressTracker(
        [string]$a,
        [string]$t
    ) {
        $this.Activity = $a
        $this.Total = $t
    }

    [bool]ExceededEstimate() { 
        return $this.Elapsed() -gt $this.Total
    }

    [int]ElapsedSince([DateTime]$since) {
        return (New-TimeSpan $since $this.LastUpdateTime).TotalSeconds
    }

    [int]Elapsed() {
        return $this.ElapsedSince($this.StartTime)
    }
}

function New-TrackedProgress {
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $Activity,

        [Parameter(Mandatory = $true)]
        [int]$Total,

        [Parameter(Mandatory = $true)]
        [string]
        $CurrentOperation
    )
    $Tracker = [ProgressTracker]::new($Activity, $Total)
    $Tracker.StartTime = $Tracker.LastUpdateTime = Get-Date
    Write-Progress -Activity $Tracker.Activity -PercentComplete 0 -CurrentOperation $CurrentOperation
    $Tracker
}

function Update-TrackedProgress {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateScript( { $null -ne $_.StartTime })]
        [ProgressTracker] $Tracker,

        [string] $CurrentOperation = $Tracker.CurrentOperation,

        [string] $Status = $Tracker.Status
    )
    
    $Tracker.CurrentOperation = $CurrentOperation
    $Tracker.Status = $Status
    $Tracker.LastUpdateTime = (Get-Date)
    $elapsed = (New-TimeSpan $Tracker.StartTime $Tracker.LastUpdateTime).TotalSeconds
    $percent = 100 * $elapsed / [Math]::Max($elapsed, $Tracker.Total)

    $currentOperationParameter = if (![string]::IsNullOrEmpty($CurrentOperation)) { "-CurrentOperation ""$CurrentOperation""" } else { "" }
    $statusParameter = if (![string]::IsNullOrEmpty($Status)) { "-Status ""$Status""" } else { "" }
    Invoke-Expression "Write-Progress -Activity `$Tracker.Activity -PercentComplete $percent $currentOperationParameter $statusParameter"
}

function Start-TrackedSleep {
    Param(
        [Parameter(Mandatory = $true)]
        [ProgressTracker] $Tracker,

        [string] $CurrentOperation = $Tracker.CurrentOperation,

        [Parameter(Mandatory = $true)]
        [ValidateScript( { $_ -ge 0 })]
        [int] $Seconds,

        [ValidateScript( { $_ -gt 0 })]
        [int] $UpdateInterval = 1
    )

    if ($Seconds -eq 0) {
        return
    }

    $sleepStart = Get-Date
    Update-TrackedProgress -Tracker $ProgressTracker -CurrentOperation $CurrentOperation
    while ((New-TimeSpan $sleepStart $Tracker.LastUpdateTime).TotalSeconds -lt $Seconds) {
        Update-TrackedProgress -Tracker $ProgressTracker -CurrentOperation $CurrentOperation
        Start-Sleep $UpdateInterval
    } 
}

function Start-TrackedRetryLoop {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ProgressTracker] $Tracker,

        [Parameter(Mandatory = $true)]
        [Scriptblock] $Action,

        [ValidateScript( { $_ -gt 0 })]
        [int] $UpdateInterval = 1
    )

    for ($iteration = 1; &$Action -ne $true; $iteration++) {
        Update-TrackedProgress $Tracker -Status "Retry attempt $iteration"
        if ($Tracker.ExceededEstimate()) {
            throw "Timed out."
        }
    }

    # Reset the status
    $Tracker.Status = $null
    Update-TrackedProgress $Tracker
}

function Get-Wlan {
    Param(
        [Parameter(Position = 0)]
        [string] $Name,
        [int] $ScanWaitTime = 3
    )

    $progressTracker = New-TrackedProgress -Activity "Getting wi-fi networks" -Total $ScanWaitTime -CurrentOperation "Requesting network scan.."
    getWlan -Name $Name -ScanWaitTime $ScanWaitTime -ProgressTracker $progressTracker
}

function Connect-Wlan {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Name,
        [ValidateRange(-1, 600)]
        [int] $Timeout = 120
    )

    $progressTracker = New-TrackedProgress -Activity "Connecting to $Name" -Total $Timeout -CurrentOperation "Looking for network.."
    
    $getWlanAction = { !(getWlan -Name $Name -ProgressTracker $progressTracker -ScanWaitTime 2) }
    Start-TrackedRetryLoop $progressTracker -Action $getWlanAction

    Update-TrackedProgress $progressTracker -CurrentOperation "Connecting..."
    (netsh wlan connect $Name) | Out-Null
 
    Update-TrackedProgress $progressTracker -CurrentOperation "Verifying connection..."
    $verifyConnectionAction = { (netsh wlan show interface | select-string -Pattern "State\s*: connected", "SSID\s*: $Name").Count -eq 2 }
    Start-TrackedRetryLoop $progressTracker -Action $verifyConnectionAction
}

function getWlan {
    Param(
        [Parameter(Position = 0)]
        [string] $Name,
        [Parameter(Mandatory = $true)]
        [int] $ScanWaitTime,
        [Parameter(Mandatory = $true)]
        [ProgressTracker] $ProgressTracker
    )

    scanWifi
    Start-TrackedSleep -Tracker $ProgressTracker -Seconds $ScanWaitTime -CurrentOperation "Waiting for network scan.."
    Update-TrackedProgress -Tracker $ProgressTracker -CurrentOperation "Reading network list.."

    $networks = (netsh wlan show networks) | ConvertFrom-String -TemplateFile "$PSScriptRoot\wlan_show_networks.txt" -UpdateTemplate
    $networks | ForEach-Object { $_.PSObject.TypeNames.Insert(0, "WifiNetwork") }
    if ($Name) {
        $networks | Where-Object { $_.Name -eq $Name }
    }
    else {
        $networks
    }
}

function scanWifi {
    $negVersion = 0
    $wlanhandle = [IntPtr]::Zero
    [Guid]$interfaceGuid = (Get-NetAdapter -Name "Wi-FI").interfaceguid
    ThrowIfFailed ([WlanAPI]::WLanOpenHandle(2, [IntPtr]::Zero, [ref] $negVersion, [ref] $wlanhandle)) "wlanapi.dll:WLanOpenHandle"
    ThrowIfFailed ([WlanAPI]::WlanScan($wlanhandle, [ref]$interfaceGuid, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)) "wlanapi.dll:WlanScan"
    ThrowIfFailed ([WlanAPI]::WlanCloseHandle($wlanhandle, [IntPtr]::Zero)) "wlanapi.dll:WlanCloseHandle"
}

function ThrowIfFailed([Int32]$hresult, [string]$api) {
    if ($hresult) {
        throw "Failed to call native API $api. HRESULT: $hresult" 
    }
}

Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
public static class WlanAPI 
{
    [DllImport("wlanapi.dll")]
    public static extern Int32 WlanOpenHandle(
        [In] UInt32 clientVersion, 
        [In, Out] IntPtr pReserved, 
        [Out] out UInt32 negotiatedVersion, 
        [Out] out IntPtr clientHandle 
    );
    [DllImport("wlanapi.dll")]
    public static extern Int32 WlanScan(
        [In] IntPtr hClientHandle,
        [In] ref Guid pInterfaceGuid,
        [In, Out] IntPtr pDot11Ssid,
        [In, Out] IntPtr pIeData,
        [In, Out] IntPtr pReserved
    );
    [DllImport("wlanapi.dll")]
    public static extern Int32 WlanCloseHandle(
        [In] IntPtr ClientHandle, 
        [In, Out] IntPtr pReserved 
    );
}
"@

Update-FormatData -AppendPath "$PSScriptRoot\formats.ps1xml"

Export-ModuleMember -Function "Get-Wlan", "Connect-Wlan"