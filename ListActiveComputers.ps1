# 
# Name : ListActiveComputers.ps1
# Purpose: Get active computer accounts from active directory by 
# checking the last logon date. Get the properties of computer
# account (name,OS,OSverion,lastlogondate and CanonicalName)
# and save it to ActiveComputers.csv file.
#

Import-Module ActiveDirectory

# get today's date
$today = Get-Date

#Get today - 60 days (2 month old)
$cutoffdate = $today.AddDays(-90)

#Get the computer accounts filtered by lastlogondate.
# Select only required properties of the computer account
# and export it to a file
Get-ADComputer -SearchBase "OU=Ozitem,OU=PC,OU=Bondy,DC=bh,DC=dom" -Properties * -Filter {LastLogonDate -gt $cutoffdate} `
| Select Name,OperatingSystem,OperatingSystemVersion, `
LastLogonDate,CanonicalName | Export-Csv  ./ActiveComputers.csv -NoTypeInformation
