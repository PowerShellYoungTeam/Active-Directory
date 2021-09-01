# Import AD module
Import-Module ActiveDirectory

# Input output files names and paths
$inputUserFile = "C:\temp\Posh_Inputs\ADGroupUserRemoval_Users.csv" # Path to input User file
$inputGroupFile = "C:\temp\Posh_Inputs\ADGroupUserRemoval_Groups.csv" # Path to input Group file
$Outfile = "C:\temp\Posh_Outputs\ADGroupUserRemoval_$(get-date -f yyyy-MM-dd-HH)_Success.csv"
$ErrorLog = "C:\temp\Posh_Outputs\ADGroupUserRemoval_$(get-date -f yyyy-MM-dd-HH)_Errorlog.log"
$Domain = "POSHYT"

# initialise error counter 
$ErrorCount=0
$Adgroups = $NULL

# initialise array for adgroups
$Adgroups =@()

# Input ADgroup file and store in Array
$Adgroups = (Import-CSV $inputGroupFile -Header ADGroup)

# Import User CSV file and loop through
Import-CSV $inputUserFile -Header User | Foreach-Object{

    # Clear variables from last loop
    $User = $Userinfo = $ErrorMessage = $null # clear variables

    # copy User into variable
    $User = $_.User

    #Search AD for User
    Try{
        $Userinfo = Get-ADUser $User -Properties * -server $Domain -ErrorAction Stop
    }catch{

        #increase error counter if something not right
        $ErrorCount += 1
        #print error message and enter into error log
        $ErrorMessage = $_.Exception.Message
        "Issue with User $($User) because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
        Write-host -ForegroundColor RED "Issue with User $($User) because $($ErrorMessage)"
    }

    # if A user was found in AD
    if($null -ne $Userinfo){
        # for each ADgroup in ADGroups Array check AD group exists and store object
        foreach($ADGroup in $ADgroups){

            #Clear variables from last loop
            $ADgroupinfo = $MembershipCheck = $ErrorMessage = $null

            # Search AD for AD group
            try{
                $ADgroupinfo = get-ADGroup $ADgroup.ADGroup  -server $Domain -ErrorAction Stop
            }catch{

                #increase error counter if something not right
                $ErrorCount += 1
                #print error message and enter into error log
                $ErrorMessage = $_.Exception.Message
                "Issue with AD Group $($ADGroup.ADGROUP) because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
                Write-host -ForegroundColor RED "Issue with AD Group $($ADGroup) because $($ErrorMessage)"
            }# end of try..catch

            # if ADgroupinfo isn't null
            if($null -ne $ADgroupinfo){

                #Check if Group has user as a member
                try{
                    $MembershipCheck = (Get-ADGroupMember -Identity $ADgroupinfo.name -Server $Domain -ErrorAction stop | Where-Object -Property name -eq $Userinfo.Name)
                }catch{
                
                    #increase error counter if something not right
                    $ErrorCount += 1
                    #print error message and enter into error log
                    $ErrorMessage = $_.Exception.Message
                    "Issue with searching AD Group $($ADGroup.ADGROUP) for $($user) because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
                    Write-host -ForegroundColor RED "Issue with searching AD Group $($ADGroup.ADGROUP) for $($user) because $($ErrorMessage)"    
                }#end of try..catch

                # if user is found in AD group
                if($null -ne $MembershipCheck){
                
                    Try{ # remember to remove whatif when actually running it -WhatIf

                        #Remove user, update console, and create and input log entry
                        Remove-ADGroupMember -Identity $ADgroupinfo.name -Members $Userinfo  -Server $Domain -ErrorAction Stop  -Confirm:$false -WhatIf
                        Write-host -ForegroundColor Green "AD User $($Userinfo.name) , $($Userinfo.Displayname) successfully removed from AD group $($ADgroupinfo.Name)"
                        [pscustomobject][ordered] @{
                            Username = $Userinfo.name
                            User = $Userinfo.Displayname
                            "Email Address" = $Userinfo.EmailAddress
                            Group = $ADgroupinfo.name
                        } | Export-csv -Path $outfile -NoTypeInformation -Append -Force
                    }catch{

                        #increase error counter if something not right
                        $ErrorCount += 1
                        #print error message and enter into error log
                        $ErrorMessage = $_.Exception.Message
                        "Issue with removing $($User) from AD Group $($ADGroup) because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
                        Write-host -ForegroundColor RED "Issue with removing $($User) from AD Group $($ADGroup) because $($ErrorMessage)"
                    } #end of try ..catch
                } # end of if membershipcheck not NULL
            } # end of if ADgroupinfo not NULL
        }#End of Foreach AD group
    } # End of if $userinfo -ne null
}# end of foreach-object

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
