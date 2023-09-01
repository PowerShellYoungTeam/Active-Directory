<#

PYT AD Reports GUI

By Steven Wight 08/08/2023

#>

################################
# Main Variables & Module Import
################################
Import-Module PYT_AD_Reports

#For Guis
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#######################
# End of Main Variables
#######################

###########
# Functions
###########

##################
# End of Functions
##################

#############
# Main Script
#############

### PYT AD Report Controller ###
$ADReportControllerForm = New-Object System.Windows.Forms.Form
$ADReportControllerForm.Text = 'PYT AD Reports'
$ADReportControllerForm.Size = New-Object System.Drawing.Size(400,270)
$ADReportControllerForm.StartPosition = 'CenterScreen'
$ADReportControllerForm.Font = ‘Microsoft Sans Serif,10’

$TextHelp = New-Object System.Windows.Forms.Label
$TextHelp.Location = New-Object System.Drawing.Point(50,20)
$TextHelp.Size = New-Object System.Drawing.Size(350,20)
$TextHelp.Text = 'Select the reports you want to run, then click RUN'
$ADReportControllerForm.Controls.Add($TextHelp)

$AllADCompCheckbox = New-Object System.Windows.Forms.Checkbox 
$AllADCompCheckbox.Location = New-Object System.Drawing.Size(10,50) 
$AllADCompCheckbox.Size = New-Object System.Drawing.Size(500,20)
$AllADCompCheckbox.Text = "Run All PYT AD Computers Report"
$AllADCompCheckbox.TabIndex = 0
$AllADCompCheckbox.Checked = $true
$ADReportControllerForm.Controls.Add($AllADCompCheckbox)

$AllADUserCheckbox = New-Object System.Windows.Forms.Checkbox 
$AllADUserCheckbox.Location = New-Object System.Drawing.Size(10,80) 
$AllADUserCheckbox.Size = New-Object System.Drawing.Size(500,20)
$AllADUserCheckbox.Text = "Run All PYT AD Users Report"
$AllADUserCheckbox.TabIndex = 1
$AllADUserCheckbox.Checked = $true
$ADReportControllerForm.Controls.Add($AllADUserCheckbox)

$PoolVDICheckbox = New-Object System.Windows.Forms.Checkbox 
$PoolVDICheckbox.Location = New-Object System.Drawing.Size(10,110) 
$PoolVDICheckbox.Size = New-Object System.Drawing.Size(500,20)
$PoolVDICheckbox.Text = "Run Pool VDI AD Group Membership Report"
$PoolVDICheckbox.TabIndex = 2
$PoolVDICheckbox.Checked = $true
$ADReportControllerForm.Controls.Add($PoolVDICheckbox)

$SecurityGroupCheckbox = New-Object System.Windows.Forms.Checkbox 
$SecurityGroupCheckbox.Location = New-Object System.Drawing.Size(10,140) 
$SecurityGroupCheckbox.Size = New-Object System.Drawing.Size(500,20)
$SecurityGroupCheckbox.Text = "Run Security Group Membership Report"
$SecurityGroupCheckbox.TabIndex = 3
$SecurityGroupCheckbox.Checked = $true
$ADReportControllerForm.Controls.Add($SecurityGroupCheckbox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,180)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'RUN'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$okButton.TabIndex = 4
$ADReportControllerForm.AcceptButton = $okButton
$ADReportControllerForm.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(200,180)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.TabIndex = 5
$ADReportControllerForm.CancelButton = $cancelButton
$ADReportControllerForm.Controls.Add($cancelButton)

$ADReportControllerForm.Topmost = $true

$result = $ADReportControllerForm.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{

    if($true -eq $AllADCompCheckbox.Checked){

        Get-PYTALLADComputers -Verbose

    }

    if($true -eq $AllADUserCheckbox.Checked){

        Get-PYTALLADusers -Verbose

    }

    if($true -eq $PoolVDICheckbox.Checked){

        Get-PYTUserGroupMembership VDI_POOL -Verbose

    }

    if($true -eq $SecurityGroupCheckbox.Checked){

        Get-PYTAdminGroups -Verbose

    }

}else{

    Write-Warning "Exiting Script"
    break

}

####################
# End of Main Script
####################az