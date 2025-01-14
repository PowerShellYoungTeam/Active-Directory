param (
    [string]$Domain,
    [string]$GroupListFilePath,
    [string]$OutputDirectory
)

# Import the Get-NestedADGroupMembers function
. ./Get-NestedADGroupMembers.ps1

# Read the list of AD group names from the file
$groupNames = Get-Content -Path $GroupListFilePath

# Initialize an array to hold the results
$results = @()

# Process each group name
foreach ($groupName in $groupNames) {
    $results += Get-NestedADGroupMembers -GroupName $groupName -Domain $Domain
}

# Get the current Unix timestamp
$timestamp = [int][double]::Parse((Get-Date -UFormat %s))

# Define the output file path
$outputFilePath = Join-Path -Path $OutputDirectory -ChildPath "ADGroupMembers_$timestamp.csv"

# Export the results to a CSV file
$results | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Exported results to $outputFilePath"

#usage
#.\ControllerScript.ps1 -Domain "YourDomain" -GroupListFilePath "C:\path\to\groupList.txt" -OutputDirectory "C:\path\to\output"