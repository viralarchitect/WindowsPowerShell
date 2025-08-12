<# 
.SYNOPSIS
  Validates a list of servers from .\servers.txt and writes results to a CSV on the Desktop:
  {LogonDomain}_Server-Validation_{YYYY-MM-DDTHH-mm-ssZ}.csv

.EXPECTED COLUMNS
  Hostname, IPv4 Address, Logon Domain, Operating System, Drives, Messages
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#--- Helpers ---
function Test-Port {
    param(
        [Parameter(Mandatory)] [string] $Computer,
        [Parameter(Mandatory)] [int] $Port,
        [int] $TimeoutMs = 2000
    )
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Computer, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $client.Close()
            return $false
        }
        $client.EndConnect($iar) | Out-Null
        $client.Close()
        return $true
    } catch {
        return $false
    }
}

function Resolve-IPv4 {
    param([Parameter(Mandatory)][string] $Host)
    try {
        # Prefer Resolve-DnsName if available; fallback to .NET
        if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
            ($res = Resolve-DnsName -Name $Host -Type A -ErrorAction Stop) | Out-Null
            return ($res | Where-Object {$_.Type -eq 'A'} | Select-Object -ExpandProperty IPAddress -First 1)
        } else {
            return [System.Net.Dns]::GetHostAddresses($Host) | Where-Object {$_.AddressFamily -eq 'InterNetwork'} | Select-Object -ExpandProperty IPAddressToString -First 1
        }
    } catch {
        return $null
    }
}

function Reverse-LookupHost {
    param([Parameter(Mandatory)][string] $IPv4)
    try {
        # Try Reverse DNS; if not present, just return the IP
        return ([System.Net.Dns]::GetHostEntry($IPv4)).HostName
    } catch {
        return $null
    }
}

#--- Setup ---
$domain   = $env:USERDOMAIN
$timestampUtc = [datetime]::UtcNow.ToString("yyyy-MM-ddTHH-mm-ssZ")
$desktop  = [Environment]::GetFolderPath('Desktop')
$outFile  = Join-Path $desktop ("{0}_Server-Validation_{1}.csv" -f $domain, $timestampUtc)

$serverListPath = Join-Path (Get-Location) 'servers.txt'
if (-not (Test-Path $serverListPath)) {
    Write-Error "servers.txt not found in $(Get-Location)."
    exit 1
}

# Quick, sane IPv4 pattern (not validating 0-255 rigorously—good enough for triage)
$ipv4Regex = '^(?:\d{1,3}\.){3}\d{1,3}$'

$results = New-Object System.Collections.Generic.List[Object]

#--- Process each server ---
Get-Content -LiteralPath $serverListPath | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object {
    $inputToken = $_.Trim()
    $row = [ordered]@{
        'Hostname'       = $null
        'IPv4 Address'   = $null
        'Logon Domain'   = $domain
        'Operating System'= $null
        'Drives'         = $null
        'Messages'       = $null
    }
    $messages = New-Object System.Collections.Generic.List[string]

    $isIp = $inputToken -match $ipv4Regex
    $hostname = $null
    $ip = $null

    try {
        if ($isIp) {
            $ip = $inputToken
            $hostname = Reverse-LookupHost -IPv4 $ip
            if (-not $hostname) { $messages.Add("Reverse DNS not found for IP $ip.") }
        } else {
            $hostname = $inputToken
            $ip = Resolve-IPv4 -Host $hostname
            if (-not $ip) {
                $row['Hostname'] = $hostname
                $messages.Add("DNS A record not found for hostname '$hostname'. Server may be offline or decommissioned.")
                throw [System.Exception]::new("No IPv4 for $hostname")
            }
        }

        # Record what we have so far
        if ($hostname) { $row['Hostname'] = $hostname }
        if ($ip)       { $row['IPv4 Address'] = $ip }

        # ICMP ping (IPv4, 1 echo)
        $pingOk = Test-Connection -ComputerName ($ip ?? $hostname) -Count 1 -Quiet -IPv4 -ErrorAction SilentlyContinue
        if (-not $pingOk) {
            $messages.Add("No ping response from $($hostname ?? $ip).")
            throw [System.Exception]::new("Ping failed")
        }

        # Port checks
        $rdpOpen = Test-Port -Computer ($ip ?? $hostname) -Port 3389
        if (-not $rdpOpen) {
            # Check SSH as an alternative (maybe Linux/Unix or appliance)
            $sshOpen = Test-Port -Computer ($ip ?? $hostname) -Port 22
            if ($sshOpen) {
                $messages.Add("RDP (3389) closed, SSH (22) open. Likely non-Windows host.")
            } else {
                $messages.Add("RDP (3389) closed and SSH (22) closed or filtered.")
            }
        } else {
            # RDP open -> try WMI/CIM
            try {
                $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName ($hostname ?? $ip) -ErrorAction Stop
                $row['Operating System'] = ($os.Caption, $os.Version -join ' ')
            } catch {
                $messages.Add("Failed to query Win32_OperatingSystem via CIM/WMI: $($_.Exception.Message)")
            }

            try {
                $drives = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName ($hostname ?? $ip) -Filter "DriveType=3" -ErrorAction Stop |
                          Select-Object -ExpandProperty DeviceID
                if ($drives) {
                    # Format like: C:\; D:\; E:\
                    $row['Drives'] = ($drives | ForEach-Object { "$_\" }) -join '; '
                } else {
                    $row['Drives'] = $null
                }
            } catch {
                $messages.Add("Failed to query fixed disks via CIM/WMI: $($_.Exception.Message)")
            }
        }

    } catch {
        # We already stash meaningful messages as we go; add the top-level error as well
        if ($_.Exception.Message) { $messages.Add("Error: $($_.Exception.Message)") }
    }

    if ($messages.Count -gt 0) { $row['Messages'] = ($messages -join ' | ') }

    # Make sure at least one of hostname/IP shows up so the failure is attributable
    if (-not $row['Hostname'] -and -not $row['IPv4 Address']) {
        if ($isIp) { $row['IPv4 Address'] = $inputToken } else { $row['Hostname'] = $inputToken }
    }

    $results.Add([pscustomobject]$row) | Out-Null
}

#--- Output ---
$results | Select-Object 'Hostname','IPv4 Address','Logon Domain','Operating System','Drives','Messages' |
    Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done. Results: $outFile"
