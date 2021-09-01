# Active-Directory
Scripts for working with AD

Remove-UsersfromListofASDgroups
A script made to drop migrated users from file shares. It takes a csv of usernames, a csv of AD group names, and searches for the users in the AD groups and removes them. it has a -whatif that needs to be removed before it will actually remove the users from the AD groups
