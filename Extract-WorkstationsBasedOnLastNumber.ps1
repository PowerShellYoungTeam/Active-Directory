#############################################################################
#
# Extract-WorkstationsBasedOnLastNumber
#
# by Steven Wight 19/11/2021
#
#############################################################################

####################################
# Initialising Variables etc..
####################################

#Grab AD Module
Import-Module ActiveDirectory

#Set Domain for AD
$Domain = "POSHYT"

#initialise Error Counter
$ErrorCount = 0

#initialise number array
$Numbers = @(9,8,7,6,5,4,3,2,1,0)

#initialise Letter array
$MachineTypeLetters = @('S','D','P','V')

#Set Naming convention filter string
$NamingConventionFilter = "POSHYTUK"

#set output path folder
$OutputFolder = "\\POSHYTUK0123S.POSHYT.corp\PowerShell_Scripts\Output"

####################################
# Functions
####################################

Function ErrorCountCheck{ # Remind me if something went wrong encase I come back with cup of tea and error has moved off console...

  If ($ErrorCount -ge 1) { #Notify if there has been errors (will be apparent if there was tho)

        Write-host -ForegroundColor Red "####################################################"
        Write-Host -ForegroundColor Red "# The script execution completed, but with errors. #"
        Write-host -ForegroundColor Red "####################################################"

        Pause

    }Else{

        Write-host -ForegroundColor Green "#############################################"
        Write-Host -ForegroundColor Green "# Script execution completed without error. #"
        Write-host -ForegroundColor Green "#############################################"
        Write-host ""

        Pause

    }

}

Function IAonError{ # if a try/catch goes man down s**tpants
    
     Write-host -ForegroundColor Red "#########"
     Write-host -ForegroundColor Red "# ERROR #"
     Write-host -ForegroundColor Red "#########"
     Write-host ""

     #Increment the error counter
     $ErrorCount += 1

     #print error message and save error in variable (Encase we want to output to error log etc...)
     $ErrorMessage = $_.Exception.Message
     Write-host -ForegroundColor Red $ErrorMessage

}

####################################
# Main Loop
####################################

foreach($Number in $Numbers) { #For each number in numbers array

    Write-Host ""
    Write-Host -ForegroundColor DarkYellow "###############################################"
    Write-Host -ForegroundColor DarkYellow "#          Extracting Number: $Number               #"
    Write-Host -ForegroundColor DarkYellow "###############################################"
    Write-Host ""

    #Clear variables for this run (you never know)
    $Machines = $outfile = $FilterString = $null
    
    #Create outfile Path
    $outfile = "$($OutputFolder)\$($Number)_Machines_$(get-date -f yyyy-MM-dd-HH-mm).csv"

    foreach($MachineTypeLetter in $MachineTypeLetters){ #For each Letter in Letters array 

        Write-Host ""
        Write-Host "Extracting Machines that end in Letter: $MachineTypeLetter"
        Write-Host ""

        #Build Filter String for Get-ADComputer
        $FilterString = "Name -like '*$($NamingConventionFilter)*' -and Name -like '*$($Number)$($MachineTypeLetter)'"

        Try{ #try and get machines for AD
            
            #Get the machines hostnames from AD as per $FilterString and Add to the $Machines array alog with the rest under this number
            $Machines += (Get-ADComputer -filter $FilterString  -server $Domain -ErrorAction Stop | Select Name  | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)

        }catch{ # if it goes pete tong while trying to get the machines from AD

           IAonError   

        } #End of Try ..Catch

    } #End of foreach Letter

    Write-Host ""
    Write-Host -ForegroundColor DarkYellow "###############################################"
    Write-Host -ForegroundColor DarkYellow "#          Extracting Number: $Number               #"
    Write-Host -ForegroundColor DarkYellow "###############################################"
    Write-Host ""

    
    Try{ #Output Hostnames for this number
    
        $Machines | Set-Content -Path $outfile -ErrorAction Stop

    }catch{ #if it goes Pete Tong creating the output file..

        IAonError   

    } #End of Try ..Catch

} #end of Foreach number

####################################
# End of Main Loop
####################################

####################################
# Check Error Count
####################################

ErrorCountCheck

####################################
# End of Script
####################################
