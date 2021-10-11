# Active-Directory
Scripts for working with AD

# Remove-UsersfromListofASDgroups

A script made to drop migrated users from file shares. It takes a csv of usernames, a csv of AD group names, and searches for the users in the AD groups and removes them. it has a -whatif that needs to be removed before it will actually remove the users from the AD groups

# Migrate-UsersBetweenADGroups

A script knocked up to migrate users between AD Groups on the same domain, it will output a error log and an output of both groups membership afterwards (This is just for a record, might do an actual coparison and output the difference (done it before but had a tight deadline for this)
