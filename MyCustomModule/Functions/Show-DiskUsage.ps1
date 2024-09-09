function Show-DiskUsage {
    # Start a new PowerShell window
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoExit -Command `"& {param(\$startTime) `" +
        "`$initialData = Get-WmiObject -Class Win32_LogicalDisk -Filter `"DriveType=3`" | ForEach-Object { `" +
        "`[pscustomobject]@{ DeviceID = `$_.DeviceID; PercentFree = (`$_.FreeSpace / `$_.Size) * 100 }`"}; `" +
        "`$global:percentFreeData = @{}; `" +
        "foreach (`$drive in `$initialData) { `" +
        "`$global:percentFreeData[`$drive.DeviceID] = [math]::round(`$drive.PercentFree, 2); `" +
        "}; `" +
        "while ($true) { `" +
            "`$diskData = Get-WmiObject -Class Win32_LogicalDisk -Filter `"DriveType=3`" | Select-Object DeviceID, Size, FreeSpace; `" +
            "`$table = `$diskData | ForEach-Object { `" +
                "`$percentFree = (`$_.FreeSpace / `$_.Size) * 100; `" +
                "`$usedSpace = `$_.Size - `$_.FreeSpace; `" +
                "`$sizeGB = [math]::round(`$_.Size / 1GB, 2); `" +
                "`$usedGB = [math]::round(`$usedSpace / 1GB, 2); `" +
                "`$freeGB = [math]::round(`$_.FreeSpace / 1GB, 2); `" +
                "`$percentFree = [math]::round(`$percentFree, 2); `" +
                "`$initialPercentFree = `$percentFreeData[`$_.DeviceID]; `" +
                "`$percentDiff = [math]::round(`$percentFree - `$initialPercentFree, 2); `" +
                "`[pscustomobject]@{ Drive=`$_.DeviceID; SizeGB=`$sizeGB; UsedGB=`$usedGB; FreeGB=`$freeGB; PercentFree=`$percentFree; PercentDiff=`$percentDiff } `" +
            "}; `" +
            "`$table | Format-Table -AutoSize | Out-String; `" +
            "Clear-Host; `" +
            "Write-Output `$table; `" +
            "Start-Sleep -Seconds 2; `" +
        "}`"` -ArgumentList (Get-Date)"

    $process = [System.Diagnostics.Process]::Start($startInfo)
    
    # Return the process object
    return $process
}
