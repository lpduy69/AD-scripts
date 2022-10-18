#Import the module needed to query ActiveDirectory
Import-Module ActiveDirectory

#Creating a table
$Tableau = @()

$ServerAD = "BOHSRVAD01"

#Retrieve groups that are present in an OU starting with Admin and in the AD domain
$Groupes = Get-ADGroup -Server $ServerAD -properties members | Where-Object {Get-ADOrganizationalUnit -Filter {name -like "Admin*"} -Server $ServerAD -SearchBase "DC=ad,DC=BondyHabitat,DC=fr" -SearchScope Subtree}

#Interrogates each group listed to find out the users
foreach ($Groupe in $Groupes) {
foreach ($Utilisateurs in $Utilisateur){ #A nested loop in the 1st queries member users.
if ($Utilisateurs.ObjectClass -eq "user"){ #A filter is applied to search only users and not any groups or computers.
$Utilisateurs = get-aduser -Server $ServerAD -identity $Utilisateurs.SamAccountName -properties Name, DisplayName, Department, Description 

#Tabulation of results
$Obj = New-Object psObject
$Obj | Add-Member -Name "Groupes" -membertype Noteproperty -Value $Groupe.SamAccountName
$Obj | Add-Member -Name "Matricule utilisateur" -membertype Noteproperty -Value $Utilisateurs.Name
$Obj | Add-Member -Name "Nom utilisateur" -membertype Noteproperty -Value $Utilisateurs.DisplayName
$Obj | Add-Member -Name "Service utilisateur" -membertype Noteproperty -Value $Utilisateurs.department
$Obj | Add-Member -Name "Description utilisateur" -membertype Noteproperty -Value $Utilisateurs.Description
$Tableau += $Obj

#Export info to CSV
echo $Tableau | Export-Csv 'C:\Temp\Export2.csv' -Delimiter ";" -Encoding "unicode" -NoTypeInformation}}}
