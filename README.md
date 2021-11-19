# Active-Directory
Scripts for working with AD

# Remove-UsersfromListofASDgroups

A script made to drop migrated users from file shares. It takes a csv of usernames, a csv of AD group names, and searches for the users in the AD groups and removes them. it has a -whatif that needs to be removed before it will actually remove the users from the AD groups

# Migrate-UsersBetweenADGroups

A script knocked up to migrate users between AD Groups on the same domain, it will output a error log and an output of both groups membership afterwards (This is just for a record, might do an actual coparison and output the difference (done it before but had a tight deadline for this)

# Extract-WorkstationsBasedOnLastNumber

A script knocked up to pull out machines (Computer Objects in AD). It's mainly used for spliting up machines for phased deployments/actions (Where there isn't better data/way to split, your just lazy or in a rush) A lot of places I worked have had machine naming conventions that have been like:- 

<CompanyNameEtc><Numbers><Letter depending on machine type>
  
Example: POSHYTUK0123S
  
So this is a PowerShell Young Team UK machine number 123 and is a Server. (For People that have different convention at their org/environment, hopefully this is a guide) 
  
So $Number should need to change unless you are only interested in a particular group.
  
$MachineTypeLetters = Edit this as needed to suit your org/enviroments, it will probably be different!
  
$NamingConventionFilter = Again Edit as needed to suit your org/enviroment
  
$OutputFolder = Obvs :)
  
I am just sticking this up in this format as if I don't do it now, I never will, however can be worked on and also thinking of turning into a Function and adding to Software Onboarding tools module.....
