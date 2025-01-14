function Get-NestedADGroupMembers {
    param (
        [string]$GroupName,
        [string]$Domain
    )

    function Get-Members {
        param (
            [string]$GroupName,
            [string]$ParentGroupName
        )

        $group = Get-ADGroup -Identity $GroupName -Server $Domain -Properties Members, GroupScope
        $members = @()

        foreach ($member in $group.Members) {
            $adObject = Get-ADObject -Identity $member -Server $Domain -Properties ObjectClass
            if ($adObject.ObjectClass -eq 'group') {
                $members += Get-Members -GroupName $adObject.Name -ParentGroupName $GroupName
            } else {
                $members += [PSCustomObject]@{
                    ParentGroupName = $ParentGroupName
                    GroupName       = $GroupName
                    GroupScope      = $group.GroupScope
                    MemberName      = $adObject.Name
                }
            }
        }

        return $members
    }

    return Get-Members -GroupName $GroupName -ParentGroupName $null
}

# Example usage:
# $result = Get-NestedADGroupMembers -GroupName "YourGroupName" -Domain "YourDomain"
# $result | Format-Table -AutoSize