Function Get-PYTALLADComputers {
    <#
    .SYNOPSIS
    Function to Extract all AD Computer objects for a domain into a spreadsheet
    by Steven Wight
    .DESCRIPTION
    Get-PYTALLADComputers <Domain> <OutputFilePath>
    .PARAMETER Domain
    Specifies the Domain you want to extract all AD Computers from.
    .PARAMETER OutputFilePath
    Specifies the path you want out file created.
    .INPUTS
    Takes strings with Domain and OutputFilePath.
    .OUTPUTS
    CSV and XLSX.
    .EXAMPLE
    Get-PYTALLADComputers PYT
    .EXAMPLE
    Get-PYTALLADComputers PYT C:\Temp
    .NOTES
    Owner - Steven Wight
    07/08/2023 - v1.0 Script creation
    Requires ActiveDirectory and ImportExcel Modules
    #>
    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a Domain to extact Computer AD Objects from, accepted values: PYT")]
        [ValidateSet('PYT')] [String] $Domain = 'PYT', 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a path where the output csv and xlsx files will go")]
        [ValidateNotNullOrEmpty()] [String] $OutputFilePath = 'c:\temp\posh_outputs'
    )

    $OutCSVFile = "$($OutputFilePath)\$($Domain)_ADcomputers_$(get-date -f yyyyMMdd).CSV"
    Write-Verbose "Output CSV file: $($OutCSVFile)"
    $OutXLSXFile = "$($OutputFilePath)\$($Domain)_ADcomputers_$(get-date -f yyyyMMdd).XLSX"
    Write-Verbose "Output XLSX file: $($OutXLSXFile)"

    try{

        Write-Verbose "Extracting all AD Computer Objects from $($Domain) Domain"
        $computers = (Get-ADComputer -filter * -server $Domain -properties Name,DNSHostname,enabled,description,OperatingSystem,serialNumber,CanonicalName,PrimaryGroup,LastLogonDate,Location,serialNumber,whenChanged,whenCreated,Modified | Select-Object Name,DNSHostname,enabled,description,OperatingSystem,CanonicalName,PrimaryGroup,LastLogonDate,Location,serialNumber,whenChanged,whenCreated) 

        # Create an object for each machine basically so we can turn the serial number into a string for export
        foreach($computer in $computers){

            [PSCustomObject][ordered]@{
                Name = $computer.Name
                DNSHostname = $computer.DNSHostname
                enabled = $computer.enabled
                description = $computer.description
                OperatingSystem = $computer.OperatingSystem
                CanonicalName = $computer.CanonicalName
                PrimaryGroup = $computer.PrimaryGroup
                LastLogonDate = $computer.LastLogonDate
                Location = $computer.Location
                serialNumber = [string]$computer.serialNumber
                whenChanged = $computer.whenChanged
                whenCreated = $computer.whenCreated
            } | Export-Csv -Path $OutCSVFile -NoTypeInformation -Append

        } 

        Write-Verbose "Creating output XLSX: $($OutXLSXFile)"
        Import-csv $OutCSVFile | Export-Excel $OutXLSXFile -AutoSize -WorksheetName "$($Domain) AD Computers" -FreezeTopRow -BoldTopRow -TableStyle Medium10 

    }Catch{

        Write-warning "ERROR: $($_.Exception.Message)" 
    
    }

} #End of Function

Function Get-PYTALLADusers {
    <#
    .SYNOPSIS
    Function to Extract all AD User objects for a domain into a spreadsheet
    by Steven Wight
    .DESCRIPTION
    Get-PYTALLADusers <Domain> <OutputFilePath>
    .PARAMETER Domain
    Specifies the Domain you want to extract all AD Users from.
    .PARAMETER OutputFilePath
    Specifies the path you want out file created.
    .INPUTS
    Takes strings with Domain and OutputFilePath.
    .OUTPUTS
    CSV and XLSX.
    .EXAMPLE
    Get-PYTALLADusers PYT
    .EXAMPLE
    Get-PYTALLADusers PYT C:\Temp
    .NOTES
    Owner - Steven Wight
    07/08/2023 - v1.0 Script creation
    Requires ActiveDirectory and ImportExcel Modules
    #>
    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a Domain to extact Users AD Objects from, accepted values: PYT")]
        [ValidateSet('PYT')] [String] $Domain = 'PYT', 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a path where the output csv and xlsx files will go")]
        [ValidateNotNullOrEmpty()] [String] $OutputFilePath = 'c:\temp\posh_outputs'
    )

    $OutCSVFile = "$($OutputFilePath)\$($Domain)_ADUsers_$(get-date -f yyyyMMdd).CSV"
    Write-Verbose "Output CSV file: $($OutCSVFile)"
    $OutXLSXFile = "$($OutputFilePath)\$($Domain)_ADUsers_$(get-date -f yyyyMMdd).XLSX"
    Write-Verbose "Output XLSX file: $($OutXLSXFile)"

    try{

        Write-Verbose "Extracting all AD User Objects from $($Domain) Domain"
        $Userinfos = (Get-ADuser -filter * -properties * -server $Domain | Select-Object name, EmployeeID, samaccountname, DisplayName, givenName, Surname, emailaddress,Description,created,Enabled,LastLogonDate, PasswordExpired,PasswordLastSet , ProfilePath ,HomeDirectory , Division, Company, Country, CanonicalName, Department , departmentNumber, Title , extensionAttribute10 , extensionAttribute4, UserPrincipalName ) 

        # Create an object for each user basically so we can turn the departmentNumber (cost centre) into a string for export
        foreach($Userinfo in $Userinfos){

            [PSCustomObject][ordered]@{
                Name = $Userinfo.Name
                EmployeeID = $Userinfo.EmployeeID
                samaccountname = $Userinfo.samaccountname
                displayname = $Userinfo.displayname
                givenName = $Userinfo.givenName
                Surname = $Userinfo.Surname
                emailaddress = $Userinfo.emailaddress
                created = $Userinfo.created
                enabled = $Userinfo.enabled
                LastLogonDate = $Userinfo.LastLogonDate
                PasswordExpired = $Userinfo.PasswordExpired
                PasswordLastSet = $Userinfo.PasswordLastSet
                ProfilePath = $Userinfo.ProfilePath
                HomeDirectory = $Userinfo.HomeDirectory
                Company = $Userinfo.Company
                Country = $Userinfo.Country
                CanonicalName = $Userinfo.CanonicalName
                Division = $Userinfo.Division
                Department = $Userinfo.Department
                departmentNumber = [string]$Userinfo.departmentNumber
                Title = $Userinfo.Title 
                extensionAttribute10 = $Userinfo.extensionAttribute10
                extensionAttribute4 = $Userinfo.extensionAttribute4
                UserPrincipalName = $Userinfo.UserPrincipalName
            } | Export-Csv -Path $OutCSVFile -NoTypeInformation -Append

        } 

        Write-Verbose "Creating output XLSX: $($OutXLSXFile)"
        Import-csv $OutCSVFile | Export-Excel $OutXLSXFile -AutoSize -WorksheetName "$($Domain) AD Users" -FreezeTopRow -BoldTopRow -TableStyle Medium10 

    }Catch{

        Write-warning "ERROR: $($_.Exception.Message)" 
    
    }

} #End of Function

function Get-PYTUserGroupMembership {
    <#
    .SYNOPSIS
    Function to extract all members of AD groups populated with user AD objects with a certain string in the names details into a spreadsheet
    by Steven Wight
    .DESCRIPTION
    Get-PYTUserGroupMembership <GroupNameSearchString> <Domain> <OutputFilePath>
    .PARAMETER GroupNameSearchString
    Specifies the string to search for AD groups we want the membership of.
    .PARAMETER Domain
    Specifies the Domain you want to extract all AD Users from.
    .PARAMETER OutputFilePath
    Specifies the path you want out file created.
    .INPUTS
    Takes strings with GroupNameSearchString, Domain and OutputFilePath. accepts pipline and arrays for GroupNameSearchString
    .OUTPUTS
    CSV and XLSX.
    .EXAMPLE
    Get-PYTUserGroupMembership VDI_POOL
    .EXAMPLE
    Get-PYTUserGroupMembership POLICY_PYT PYT
    .EXAMPLE
    Get-PYTUserGroupMembership VDI_POOL PYT C:\Temp
    .NOTES
    Owner - Steven Wight
    07/08/2023 - v1.0 Script creation
    Requires ActiveDirectory and ImportExcel Modules
    Doesn't do foreign domain accounts
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage="Enter in a search string for the AD groups name. e.g. VDI_POOL")]
        [ValidateNotNullOrEmpty()] [String] $GroupNameSearchString,
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a Domain to extact AD groups members from, accepted values: PYT")]
        [ValidateSet('PYT','PRBUK','UKGCBPRO')] [String] $Domain = 'PYT', 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a path where the output csv and xlsx files will go")]
        [ValidateNotNullOrEmpty()] [String] $OutputFilePath = 'c:\temp\posh_outputs'
    )
    
    $OutCSVFile = "$($OutputFilePath)\$($Domain)_$($GroupNameSearchString)_UserGroupMembershipReport_$(get-date -f yyyyMMdd).CSV"
    Write-Verbose "Output CSV file: $($OutCSVFile)"
    $OutXLSXFile = "$($OutputFilePath)\$($Domain)_$($GroupNameSearchString)_UserGroupMembershipReport_$(get-date -f yyyyMMdd).XLSX"
    Write-Verbose "Output XLSX file: $($OutXLSXFile)"

    try{

        Write-Verbose "Extracting all $($GroupNameSearchString) named AD groups from $($Domain) Domain"
        $Groups = (get-adgroup -Filter "name -like '*$($GroupNameSearchString)*' -and name -notlike '*zz*'" -server $Domain -ErrorAction stop | Select-Object -Expand name)

    }catch{

        $ErrorMessage = $_.Exception.Message
        Write-Warning  "Wasn't able to extract $($Domain)/$($GroupNameSearchString) AD Groups because of $($ErrorMessage)"

    }# end of try ...catch

    foreach($Group in $Groups){ 

        #clear variables from last groups so no data leakage.
        $Users = $null  
        
        Write-Verbose "Extracting all users of $($Domain)/$($Group) AD group"
        
        Try{ # extract group membership into array

            $Users = (get-adgroup $Group -properties Member -server $Domain | Select-Object -Expand Member )

            Foreach($User in $Users){# for each user in the array
    
                #clear variables
                $Userinfo = $null
    
                try{#Extract user details into object and write to CSV

                    $Userinfo =  (Get-ADUser $User -Properties Name, displayname, emailaddress, enabled, Country, Division, Company, Department, departmentNumber,Title -server $Domain)  

                    [pscustomobject][ordered] @{
                        Domain = $Domain
                        Group =  $Group
                        Name = $Userinfo.name
                        displayname = $Userinfo.displayname
                        emailaddress = $Userinfo.emailaddress
                        enabled = $Userinfo.enabled
                        Country = $Userinfo.Country
                        Division = $Userinfo.Division
                        Company = $Userinfo.Company
                        Department = $Userinfo.Department
                        departmentNumber = [string]$Userinfo.departmentNumber
                        Title = $Userinfo.Title                       
                    }|Export-csv $OutCSVFile -NoTypeInformation -Append

                }catch{

                    $ErrorMessage = $_.Exception.Message
                    Write-Warning  "Wasn't able to extract $($User) details because of $($ErrorMessage)"

                } # end of try ...catch

            } #end of foreach user loop

        }catch{

            $ErrorMessage = $_.Exception.Message
            Write-Warning  "Wasn't able to extract $($Group) AD group membership because of $($ErrorMessage)"

        }# end of try ...catch

    
    }#end of foreach group loop

    Write-Verbose "Creating output XLSX: $($OutXLSXFile)"
    Import-csv $OutCSVFile | Export-Excel $OutXLSXFile -AutoSize -WorksheetName $GroupNameSearchString -FreezeTopRow -BoldTopRow -TableStyle Medium10 

}# End of Function

function Get-PYTAdminGroups {
    <#
    .SYNOPSIS
    Function to extract all Admin Groups & memberships details into a spreadsheet, made specifically for admin groups (can be used for others)
    by Steven Wight
    .DESCRIPTION
    Get-PYTAdminGroups <GroupNameSearchString> <Domains> <OutputFilePath>
    .PARAMETER GroupNameSearchString
    Specifies the string to search for Admin AD groups we want the membership of.
    .PARAMETER Domains
    Specifies the Domains you want to extract all admin AD Users from.
    .PARAMETER OutputFilePath
    Specifies the path you want out file created.
    .INPUTS
    Takes strings with GroupNameSearchString, Domain and OutputFilePath. accepts pipline and arrays for GroupNameSearchString
    .OUTPUTS
    CSV and XLSX.
    .EXAMPLE
    Get-PYTAdminGroups 
    .EXAMPLE
    Get-PYTAdminGroups ADM DL U Admin
    .EXAMPLE
    Get-PYTAdminGroups DLG_Admin PYT C:\Temp
    .NOTES
    Owner - Steven Wight
    07/08/2023 - v1.0 Script creation
    Requires ActiveDirectory and ImportExcel Modules
    made specifically for SANUK & PRBUK
    Doesn't do foreign domain accounts
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage="Enter in a search string for the Admin's AD groups name. e.g. DLG_Admin")]
        [ValidateNotNullOrEmpty()] [String] $GroupNameSearchString = 'DLG_Admin',
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a Domains to extact AD groups members from, accepted values (multiple): PYT, PYTUAT, PYTDEV")]
        [ValidateSet('PYT','PYTUAT','PYTDEV')] [String] $Domain = ('PYT','PYTUAT','PYTDEV'), 
        [Parameter(Mandatory=$false,
        ValueFromPipeline = $false,
        ValueFromPipelineByPropertyName = $false,
        HelpMessage="Enter in a path where the output csv and xlsx files will go")]
        [ValidateNotNullOrEmpty()] [String] $OutputFilePath = 'c:\temp\posh_outputs'
    )
    
    $OutCSVFile = "$($OutputFilePath)\$($Domain)_$($GroupNameSearchString)_UserGroupMembershipReport_$(get-date -f yyyyMMdd).CSV"
    Write-Verbose "Output CSV file: $($OutCSVFile)"
    $OutXLSXFile = "$($OutputFilePath)\$($Domain)_$($GroupNameSearchString)_UserGroupMembershipReport_$(get-date -f yyyyMMdd).XLSX"
    Write-Verbose "Output XLSX file: $($OutXLSXFile)"

    foreach($Domain in $Domains){

        try{

            Write-Verbose "Extracting all $($GroupNameSearchString) named AD groups from $($Domain) Domain"
            $Groups = (get-adgroup -Filter "name -like '*$($GroupNameSearchString)*' -and name -notlike '*zz*'" -server $Domain -ErrorAction stop)

        }catch{

            $ErrorMessage = $_.Exception.Message
            Write-Warning -ForegroundColor RED "Wasn't able to extract $($Domain)/$($GroupNameSearchString) AD Groups because of $($ErrorMessage)"

        }# end of try ...catch

        foreach($Group in $Groups){ 

            #clear variables from last groups so no data leakage.
            $Users = $null  
            
            Write-Verbose "Extracting all users of $($Domain)/$($Group.Name) AD group"
            
            Try{ # extract group membership into array

                $Users = (get-adgroup $Group.Name -properties Member -server $Domain | Select-Object -Expand Member )

                Foreach($User in $Users){# for each user in the array
        
                    #clear variables
                    $Userinfo = $null
        
                    try{#Extract user details into object and write to CSV

                        $Userinfo =  (Get-ADUser $User -Properties Name, displayname, emailaddress, enabled, Country, Division, Company, Department, departmentNumber,Title -server $Domain)  

                        [pscustomobject][ordered] @{
                            Domain = $Domain
                            Group =  $Group.Name
                            OU = $Group.CanonicalName
                            Description = $Group.Description
                            GroupCategory = $Group.GroupCategory
                            GroupScope = $Group.GroupScope
                            Name = $Userinfo.name
                            DisplayName = $Userinfo.displayname
                            EmailAddress = $Userinfo.emailaddress
                            Enabled = $Userinfo.enabled
                            Country = $Userinfo.Country
                            Division = $Userinfo.Division
                            Company = $Userinfo.Company
                            Department = $Userinfo.Department
                            DepartmentNumber = [string]$Userinfo.departmentNumber
                            Title = $Userinfo.Title                       
                        }|Export-csv $OutCSVFile -NoTypeInformation -Append

                    }catch{

                        $ErrorMessage = $_.Exception.Message
                        Write-Warning  "Wasn't able to extract $($Domain)/$($User) of $($Domain)/$($Group) details because of $($ErrorMessage)"

                    } # end of try ...catch

                } #end of foreach user loop

            }catch{

                $ErrorMessage = $_.Exception.Message
                Write-Warning  "Wasn't able to extract $($Domain)/$($Group) AD group membership because of $($ErrorMessage)"

            }# end of try ...catch

        
        }#end of foreach group loop

    }# End of foreach domain loop

    Write-Verbose "Creating output XLSX: $($OutXLSXFile)"
    Import-csv $OutCSVFile | Export-Excel $OutXLSXFile -AutoSize -WorksheetName $GroupNameSearchString -FreezeTopRow -BoldTopRow -TableStyle Medium10 

}# End of Function