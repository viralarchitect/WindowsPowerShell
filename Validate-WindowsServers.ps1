
<# 
.NAME
    Windows Server Validator
.SYNOPSIS
    UI for sysadmins to quickly validate multiple windows servers and services after a reboot or service outage.

#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,400)
$Form.text                       = "Windows Server Scanner"
$Form.TopMost                    = $false

$ServerList                      = New-Object system.Windows.Forms.TextBox
$ServerList.multiline            = $true
$ServerList.text                 = "<hostname(s)>"
$ServerList.width                = 236
$ServerList.height               = 352
$ServerList.location             = New-Object System.Drawing.Point(16,29)
$ServerList.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ExecuteScan                     = New-Object system.Windows.Forms.Button
$ExecuteScan.text                = "Scan"
$ExecuteScan.width               = 122
$ExecuteScan.height              = 30
$ExecuteScan.location            = New-Object System.Drawing.Point(264,17)
$ExecuteScan.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RdpChk                          = New-Object system.Windows.Forms.CheckBox
$RdpChk.text                     = "RDP Service"
$RdpChk.AutoSize                 = $false
$RdpChk.width                    = 95
$RdpChk.height                   = 20
$RdpChk.location                 = New-Object System.Drawing.Point(274,98)
$RdpChk.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$WinRmChk                        = New-Object system.Windows.Forms.CheckBox
$WinRmChk.text                   = "WinRM Service"
$WinRmChk.AutoSize               = $false
$WinRmChk.width                  = 95
$WinRmChk.height                 = 20
$WinRmChk.location               = New-Object System.Drawing.Point(274,82)
$WinRmChk.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LastRebootChk                   = New-Object system.Windows.Forms.CheckBox
$LastRebootChk.text              = "Last Reboot"
$LastRebootChk.AutoSize          = $false
$LastRebootChk.width             = 95
$LastRebootChk.height            = 20
$LastRebootChk.location          = New-Object System.Drawing.Point(273,65)
$LastRebootChk.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$IntervalTxt                     = New-Object system.Windows.Forms.TextBox
$IntervalTxt.multiline           = $false
$IntervalTxt.text                = "5"
$IntervalTxt.width               = 27
$IntervalTxt.height              = 20
$IntervalTxt.location            = New-Object System.Drawing.Point(304,375)
$IntervalTxt.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RefreshLbl2                     = New-Object system.Windows.Forms.Label
$RefreshLbl2.text                = "seconds"
$RefreshLbl2.AutoSize            = $true
$RefreshLbl2.width               = 25
$RefreshLbl2.height              = 10
$RefreshLbl2.location            = New-Object System.Drawing.Point(343,378)
$RefreshLbl2.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RefreshLbl1                     = New-Object system.Windows.Forms.Label
$RefreshLbl1.text                = "refresh every"
$RefreshLbl1.AutoSize            = $true
$RefreshLbl1.width               = 25
$RefreshLbl1.height              = 10
$RefreshLbl1.location            = New-Object System.Drawing.Point(304,357)
$RefreshLbl1.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SvcChk                          = New-Object system.Windows.Forms.CheckBox
$SvcChk.text                     = "Other Services"
$SvcChk.AutoSize                 = $false
$SvcChk.width                    = 95
$SvcChk.height                   = 20
$SvcChk.location                 = New-Object System.Drawing.Point(273,136)
$SvcChk.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SvcListTxt                      = New-Object system.Windows.Forms.TextBox
$SvcListTxt.multiline            = $true
$SvcListTxt.text                 = "<serviceName(s)>"
$SvcListTxt.width                = 118
$SvcListTxt.height               = 193
$SvcListTxt.location             = New-Object System.Drawing.Point(273,154)
$SvcListTxt.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$serversLbl                      = New-Object system.Windows.Forms.Label
$serversLbl.text                 = "Server List:"
$serversLbl.AutoSize             = $true
$serversLbl.width                = 25
$serversLbl.height               = 10
$serversLbl.location             = New-Object System.Drawing.Point(16,11)
$serversLbl.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($ServerList,$ExecuteScan,$RdpChk,$WinRmChk,$LastRebootChk,$IntervalTxt,$RefreshLbl2,$RefreshLbl1,$SvcChk,$SvcListTxt,$serversLbl))







[void]$Form.ShowDialog()