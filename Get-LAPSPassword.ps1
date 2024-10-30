<#
============================================================================================================================
Script: Get-LAPSPassword
Author: Smart Ace Designs

Notes:
Use this script to retrieve the LAPS managed password of a member server in the corporate domain.
============================================================================================================================
#>

#region Settings
$REQUIRED_MODULES = @("LAPS")
$SUPPORT_CONTACT = "Smart Ace Designs"
#endregion

#region Assemblies
Add-Type -AssemblyName System.Windows.Forms
#endregion

#region Appearance
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Controls
$FormMain = New-Object -TypeName System.Windows.Forms.Form
$GroupBoxMain = New-Object -TypeName System.Windows.Forms.GroupBox
$LabelServerName = New-Object -TypeName System.Windows.Forms.Label
$TextBoxServerName = New-Object -TypeName System.Windows.Forms.TextBox
$LabelExpirationDate = New-Object -TypeName System.Windows.Forms.Label
$TextBoxExpirationDate = New-Object -TypeName System.Windows.Forms.TextBox
$LabelPassword = New-Object -TypeName System.Windows.Forms.Label
$TextBoxPassword = New-Object -TypeName System.Windows.Forms.TextBox
$ButtonRun = New-Object -TypeName System.Windows.Forms.Button
$ButtonClose = New-Object -TypeName System.Windows.Forms.Button
$StatusStripMain = New-Object -TypeName System.Windows.Forms.StatusStrip
$ToolStripStatusLabelMain = New-Object -TypeName System.Windows.Forms.ToolStripStatusLabel
$ErrorProviderMain = New-Object -TypeName System.Windows.Forms.ErrorProvider
#endregion

#region Forms
$ShowFormMain =
{
    $FormWidth = 330
    $FormHeight = 260

    $FormMain.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path)
    $FormMain.Text = "Get LAPS Password"
    $FormMain.ClientSize = New-Object -TypeName System.Drawing.Size($FormWidth,$FormHeight)
    $FormMain.Font = New-Object -TypeName System.Drawing.Font("MS Sans Serif",8)
    $FormMain.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $FormMain.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $FormMain.MaximizeBox = $false
    $FormMain.AcceptButton = $ButtonRun
    $FormMain.CancelButton = $ButtonClose
    $FormMain.Add_Shown($FormMain_Shown)

    $GroupBoxMain.Location = New-Object -TypeName System.Drawing.Point(10,5)
    $GroupBoxMain.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 20),($FormHeight - 80))
    $FormMain.Controls.Add($GroupBoxMain)

    $LabelServerName.Location = New-Object -TypeName System.Drawing.Point(15,15)
    $LabelServerName.AutoSize = $true
    $LabelServerName.Text = "Server Name:"
    $GroupBoxMain.Controls.Add($LabelServerName)

    $TextBoxServerName.Location = New-Object -TypeName System.Drawing.Point(15,35)
    $TextBoxServerName.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 50),20)
    $TextBoxServerName.TabIndex = 0
    $TextBoxServerName.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper
    $TextBoxServerName.MaxLength = 15
    $TextBoxServerName.Add_TextChanged($TextBoxServerName_TextChanged)
    $GroupBoxMain.Controls.Add($TextBoxServerName)

    $LabelExpirationDate.Location = New-Object -TypeName System.Drawing.Point(15,70)
    $LabelExpirationDate.AutoSize = $true
    $LabelExpirationDate.Text = "Password Expiration Date:"
    $GroupBoxMain.Controls.Add($LabelExpirationDate)

    $TextBoxExpirationDate.Location = New-Object -TypeName System.Drawing.Point(30,90)
    $TextBoxExpirationDate.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 70),20)
    $TextBoxExpirationDate.BorderStyle = "None"
    $TextBoxExpirationDate.ReadOnly = $true
    $TextBoxExpirationDate.Text = "Waiting..."
    $GroupBoxMain.Controls.Add($TextBoxExpirationDate)

    $LabelPassword.Location = New-Object -TypeName System.Drawing.Point(15,125)
    $LabelPassword.AutoSize = $true
    $LabelPassword.Text = "Administrator Password:"
    $GroupBoxMain.Controls.Add($LabelPassword)

    $TextBoxPassword.Location = New-Object -TypeName System.Drawing.Point(30,145)
    $TextBoxPassword.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 70),20)
    $TextBoxPassword.BorderStyle = "None"
    $TextBoxPassword.ReadOnly = $true
    $TextBoxPassword.Text = "Waiting..."
    $TextBoxPassword.Add_DoubleClick($TextBoxPassword_DoubleClick)
    $GroupBoxMain.Controls.Add($TextBoxPassword)

    $ButtonRun.Location = New-Object -TypeName System.Drawing.Point(($FormWidth - 175),($FormHeight - 60))
    $ButtonRun.Size = New-Object -TypeName System.Drawing.Size(75,25)
    $ButtonRun.TabIndex = 100
    $ButtonRun.Text = "Run"
    $ButtonRun.Enabled = $fase
    $ButtonRun.Add_Click($ButtonRun_Click)
    $FormMain.Controls.Add($ButtonRun)

    $ButtonClose.Location = New-Object -TypeName System.Drawing.Point(($FormWidth - 85),($FormHeight - 60))
    $ButtonClose.Size = New-Object -TypeName System.Drawing.Size(75,25)
    $ButtonClose.TabIndex = 101
    $ButtonClose.Text = "Close"
    $ButtonClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $FormMain.Controls.Add($ButtonClose)

    $StatusStripMain.SizingGrip = $false
    $StatusStripMain.Font = New-Object -TypeName System.Drawing.Font("MS Sans Serif",8)
    [void]$StatusStripMain.Items.Add($ToolStripStatusLabelMain)
    $FormMain.Controls.Add($StatusStripMain)

    [void]$FormMain.ShowDialog()
    $FormMain.Dispose()
}
#endregion

#region Handlers
$FormMain_Shown =
{
    $ToolStripStatusLabelMain.Text = "Ready"
    $StatusStripMain.Update()
    $FormMain.Activate()
}

$TextBoxServerName_TextChanged =
{
    $ToolStripStatusLabelMain.Text = "Ready"
    $StatusStripMain.Update()
    $ErrorProviderMain.Clear()
    $ButtonRun.Enabled = $true
    $TextBoxPassword.Text = "Waiting..."
    $TextBoxExpirationDate.Text = "Waiting..."
    $TextBoxPassword.Font = New-Object -TypeName System.Drawing.Font("Microsoft Sans Serif",8)

    if ($TextBoxServerName.TextLength -eq 0)
    {
        $ButtonRun.Enabled = $false
    }
    elseif ($TextBoxServerName.Text -match "[^a-z0-9A-Z\-]")
    {
        $ErrorProviderMain.SetIconPadding($TextBoxServerName,-20)
        $ErrorProviderMain.SetError($TextBoxServerName,"The server name contains an invalid character.")
        $ButtonRun.Enabled = $false
        $ToolStripStatusLabelMain.Text = "Not ready...please correct error"
        $StatusStripMain.Update()
    }
}

$TextBoxPassword_DoubleClick =
{
    Set-Clipboard -Value $TextBoxPassword.Text
    [void][System.Windows.Forms.MessageBox]::Show(
        "LAPS password copied to clipboard.`n`nPlease manually clear your clipboard contents when the password is no longer needed.",
        "Warning",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}

$ButtonRun_Click = 
{
    $ToolStripStatusLabelMain.Text = "Working...please wait"
    $FormMain.Controls | Where-Object {$PSItem -isnot [System.Windows.Forms.StatusStrip]} | ForEach-Object {$PSItem.Enabled = $false}
    $FormMain.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    try
    {
        $ServerName = ($TextBoxServerName.Text).Trim()
        if ($Results = Get-LapsADPassword -AsPlainText -Identity  $ServerName -WarningAction SilentlyContinue)
        {
            $TextBoxExpirationDate.Text = $Results.ExpirationTimeStamp
            $TextBoxPassword.Font = New-Object -TypeName System.Drawing.Font("Microsoft Sans Serif",16)
            $TextBoxPassword.Text = $Results.Password
        }
        else
        {
            $TextBoxExpirationDate.Text = "No LAPS data found..."
            $TextBoxPassword.Font = New-Object -TypeName System.Drawing.Font("Microsoft Sans Serif",8)
            $TextBoxPassword.Text = "No LAPS data found..."
        }
    }
    catch
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $PSItem.Exception.Message + "`n`nContact $SUPPORT_CONTACT for technical support.",
            "Exception",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $FormMain.Close()
    }

    $FormMain.Controls | ForEach-Object {$PSItem.Enabled = $true}
    $FormMain.ResetCursor()
    $TextBoxServerName.Focus()
    $ToolStripStatusLabelMain.Text = "Ready"
    $StatusStripMain.Update()
}
#endregion

#region Main
$MissingModules = @()
foreach ($Module in $REQUIRED_MODULES)
{
    if ($REQUIRED_MODULES -notcontains (Get-Module -ListAvailable -Name $Module | Get-Unique)) {$MissingModules += (" $([char]8226) " + $Module)}
}

if ($MissingModules)
{
    [void][System.Windows.Forms.MessageBox]::Show(
        "The following PowerShell modules are required:`n`n$($MissingModules -join "`n")`n`nPlease contact $SUPPORT_CONTACT for technical support.",
        "Requirements",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
else
{
    [void][System.Windows.Forms.MessageBox]::Show(
        "This script will display the LAPS password in plain text for a specified domain-joined computer.`n`nPlease use with caution.",
        "Warning",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    Invoke-Command -ScriptBlock $ShowFormMain
}
