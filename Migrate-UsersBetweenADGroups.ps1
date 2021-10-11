<#
#
# Migrate-UsersBetweenADGroups
#
# Migrate User objects between Groups in same domain
#
# By Steven Wight (11/10/2021)
#
#>

# Import AD module
Import-Module ActiveDirectory

# Input output files names and paths
$inputFile = "c:\Posh_inputs\ADGroups.csv" # Path to input file
$Outfile = "c:\Posh_outputs\ADGroupMigration_Users_$(get-date -f yyyy-MM-dd-HH)_Success.csv"
$ErrorLog = "c:\Posh_outputs\ADGroupMigration_Users_$(get-date -f yyyy-MM-dd-HH)_Errorlog.log"

#initialise error counter 
$ErrorCount=0

#Set Domain
$Domain = "POSHYT"

# Import CSV file and loop through
Import-CSV $InputFile -Header Source, Target | Foreach-Object{

    $Source = $Target = $SourceChecks = $TargetChecks  = $ErrorMessage = $null # clear variables

    # copy input into variables
    $Source = $_.Source
    $Target = $_.Target

    try{ 

        # Copy Membership over to Target AD Group NB: - remember to remove whatif when actually running it -WhatIf
        Get-ADGroupMember $Source -Server $Domain -ErrorAction Stop | Get-ADUser -Server $Domain  | ForEach-Object {Add-ADGroupMember -Identity $Target -Members $_ -Server $Domain -WhatIf}
    
        Write-host -ForegroundColor Green "AD Users Migrated from $($Source) to AD group $($Target)"

        #Check Source and Target membership post migration
        $SourceChecks = (Get-ADGroupMember $Source -Server $Domain  | Get-ADUser -Properties Displayname,EmailAddress -Server $Domain )
        foreach($SourceCheck in $SourceChecks){

            [pscustomobject][ordered] @{
                "Ad Group Type" = "Source"
                "AD Group" = $Source
                Username = $SourceCheck.name
                User = $SourceCheck.Displayname
                "Email Address" = $SourceCheck.EmailAddress
            } | Export-csv -Path $outfile -NoTypeInformation -Append -Force
        }#end of foreach

        $TargetChecks = (Get-ADGroupMember $Target -Server $Domain | Get-ADUser -Server $Domain)

        foreach($TargetCheck in $TargetChecks){

            [pscustomobject][ordered] @{
                "Ad Group Type" = "Target"
                "AD Group" = $Target
                Username = $TargetCheck.name
                User = $TargetCheck.Displayname
                "Email Address" = $TargetCheck.EmailAddress
            } | Export-csv -Path $outfile -NoTypeInformation -Append -Force
        }#end of foreach  
    }catch{
        #increase error counter if something not right
        $ErrorCount += 1
        #print error message
        $ErrorMessage = $_.Exception.Message
        "Issue Migrating AD Group $($Source) to $($Target) because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
        Write-host -ForegroundColor RED "Issue Migrating AD Group $($Source) to $($Target) because $($ErrorMessage)" 
    }  
}

If ($ErrorCount -ge1) {

        Write-host "-----------------"
        Write-Host -ForegroundColor Red "The script execution completed, but with errors. See $($ErrorLog)"
        Write-host "-----------------"
        Pause
}Else{
        Write-host "-----------------"
        Write-Host -ForegroundColor Green "Script execution completed without error."
        Write-host "-----------------"
        Pause
}
