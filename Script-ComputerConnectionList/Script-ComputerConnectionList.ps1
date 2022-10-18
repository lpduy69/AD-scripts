#############################################################################################
# Description: Inventaire des PdT		
# CAUTION - Required: #
# - Create directory C:\Tools-Outsourcing												#
# - Create directory C:\Tools-Outsourcing\Reports										#
# - Create directory C:\Tools-Outsourcing\Reports\"TypeDeCheck" (ici, WksListing)		#
# - Create directory C:\Tools-Outsourcing\Scripts										#
# - Create directory C:\Tools-Outsourcing\Scripts\"TypeDeCheck" (ici, WksListing)		#
#############################################################################################
#############################################################################################
# If impossible to create these directories, consider modifying $ReportFolder			#
#############################################################################################

#############################################################################################
# Definition of basic variables #
# A .CSV file listing all positions is created in $ReportFolder #
# This file must be retrieved from our filer, converted to .PDF after formatting #
# in Excel							#
#############################################################################################
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReportFolder = "$scriptPath\..\..\Reports\WksListing\"
$Date = $(Get-Date -Format "yyyy-MM-dd")
$HTMfileName = $Date + "_InventairePdT.htm"
$Report = $ReportFolder + $HTMfileName
$ADsrv = "ERGDCVP01.selpro.intra"
$Ret = (Get-Date).AddMonths(-2)
#############################################################################################

#############################################################################################
# Creation of the Reports folder if it does not exist										#
#############################################################################################
Function Create 
{
	if(!(Test-Path $ReportFolder))
	{
		New-Item $ReportFolder -type directory 
	}
}
#############################################################################################

#############################################################################################
# Define HTML report body parameters in the $Report variable. 								#
#############################################################################################
 Function Corps 
{
	New-Item $Report -type file -force
	Add-Content $Report "<html>" 
	Add-Content $Report "<head>" 
	Add-Content $Report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
	Add-Content $Report "<title>Liste des postes de travail au $Date</title>"
	Add-Content $Report '<STYLE TYPE="text/css">' 
	Add-Content $Report "<!--" 
	Add-Content $Report "td {" 
	Add-Content $Report "font-family: Tahoma;" 
	Add-Content $Report "font-size: 11px;" 
	Add-Content $Report "border-top: 1px solid #999999;" 
	Add-Content $Report "border-right: 1px solid #999999;" 
	Add-Content $Report "border-bottom: 1px solid #999999;" 
	Add-Content $Report "border-left: 1px solid #999999;" 
	Add-Content $Report "pAdding-top: 0px;" 
	Add-Content $Report "pAdding-right: 0px;" 
	Add-Content $Report "pAdding-bottom: 0px;" 
	Add-Content $Report "pAdding-left: 0px;" 
	Add-Content $Report "}" 
	Add-Content $Report "body {" 
	Add-Content $Report "margin-left: 5px;" 
	Add-Content $Report "margin-top: 5px;" 
	Add-Content $Report "margin-right: 0px;" 
	Add-Content $Report "margin-bottom: 10px;" 
	Add-Content $Report "" 
	Add-Content $Report "table {" 
	Add-Content $Report "border: thin solid #000000;" 
	Add-Content $Report "}" 
	Add-Content $Report "-->" 
	Add-Content $Report "</style>" 
	Add-Content $Report "</head>" 
	Add-Content $Report "<body>" 
	Add-Content $Report "<table width='100%'>" 
	Add-Content $Report "<tr bgcolor='#0082F1'>" 
	Add-Content $Report "<td colspan='6' height='25' align='center'>" 
	Add-Content $Report "<font face='tahoma' color='#99CEEE' size='4'><strong>Liste des postes de travail actifs dans AD au $Date</strong></font>" 
	Add-Content $Report "</td>" 
	Add-Content $Report "</tr>" 
	Add-Content $Report "</table>" 
	Add-Content $Report "<table width='100%'>"
	Add-Content $Report "<tr bgcolor='#0082F1'>" 
	Add-Content $Report "<td colspan='6' height='25' align='center'>" 
	Add-Content $Report "<font face='tahoma' color='#99CEEE' size='3'><strong>Remarque : les informations de connexions ne peuvent être récupérées que si l'ordinateur est en ligne</strong></font>"
	Add-Content $Report "</td>" 
	Add-Content $Report "</tr>" 
	Add-Content $Report "</table>" 
}
#############################################################################################

#############################################################################################
# Execution of the search command for positions in the AD with information 					#
# Search on each workstation for the last user connected with the connection date			#
#############################################################################################
Function ListWks
{
	#Import of the module needed to query ActiveDirectory
	Import-Module ActiveDirectory
	
	# Search for active workstations with information
	$WksInfo = Get-ADComputer -Server $ADsrv -Filter {(Enabled -eq "True")} -SearchBase "OU=Portable,OU=Ordinateurs,DC=selpro,DC=intra" -SearchBase "OU=Tablettes,OU=Ordinateurs,DC=selpro,DC=intra" -Properties LastLogonDate,Description,ipv4Address | Select-Object Name,ipv4Address,@{Name='LastLogonDate';Expresgetsion={$_.LastLogonDate.ToString("dd/MM/yyyy HH:mm")}},Description
	#Creation of a job list
	$WksList = New-Object System.Collections.ArrayList
	foreach($wks in $WksInfo)
	{
		$WksList.Add(($wks.Name.ToUpper().trim())) | Out-Null
	}

	# Search for user login information (requires logging into workstations)
	$WksResult = @()
	foreach($wks in $WksList)
	{
		$WksUniq = $WksInfo | Where-Object { $_.Name -eq $wks }
		$pingwks = (ping -n 1 -4 $wks)
		if("$pingwks" -NotMatch "TTL=")
		{
			$WksObj = $WksUniq | Select-Object *,@{Name="User";Expression={"Ordinateur injoignable"}},@{Name="LastUserLogon";Expression={""}}
			$WksResult += $WksObj
			Continue
		}
		else
		{
			$WinEvent = Get-WinEvent -Computername $wks -FilterHashtable @{Logname='Security';ID=4624;StartTime=$Ret} | Where-Object { ($_.Properties[4].Value -Match  "^S-1-5-21") -And ($_.Properties[5].Value -NotMatch  "owentis") } | Select-Object -first 1 @{Name='User';Expression={$_.Properties[5].Value}},@{Name="LastUserLogon";Expression={$_.TimeCreated.ToString("dd/MM/yyyy HH:mm")}}
			$userCnx = $WinEvent.User
			$dateCnx = $WinEvent.LastUserLogon
			$WksObj = $WksUniq | Select-Object *,@{Name="User";Expression={$userCnx}},@{Name="LastUserLogon";Expression={$dateCnx}}
			$WksResult += $WksObj
		}
	}
	$WksResult = $WksResult | Sort-Object -Property Name

	# Filling in the report
	Add-Content $Report "<table width='100%'>"
	Add-Content $Report "<tr bgcolor='#99CEEE'>" 
	Add-Content $Report "<td width='15%' align='center'><B>Ordinateur</B></td>"
	Add-Content $Report "<td width='15%' align='center'><B>IP (résolution DNS)</B></td>"
	Add-Content $Report "<td width='15%' align='center'><B>Dernière authentification machine (AD)</B></td>"	
	Add-Content $Report "<td width='15%' align='center'><B>Description du poste (AD)</B></td>"	
	Add-Content $Report "<td width='20%' align='center'><B>Utilisateur</B></td>"
	Add-Content $Report "<td width='20%' align='center'><B>Connecté le</B></td>"			
	Add-Content $Report "</tr>"
	
	foreach($Obj in $WksResult)
	{
		Add-Content $Report "<tr>" 
		# Test for duplicate IPs / DNS resolution issue
		$test = (($WksResult | Where-Object { $_.ipv4Address -eq $($Obj.ipv4Address) -And $_.ipv4Address -ne $Null}) | Measure-Object).Count
		# Management of duplicate and unreachable colors
		if($test -gt 1 -And $($Obj.User) -eq "Ordinateur injoignable")
		{
			$IpColor = "'#FF9B7A'"
			$StdColor = "'#B3B3B3'"
		}
		elseif($test -gt 1 -And $($Obj.User) -ne "Ordinateur injoignable")
		{
			$IpColor = "'#FF9B7A'"
			$StdColor = "'#99CEEE'"
		}
		elseif($test -le 1 -And $($Obj.User) -eq "Ordinateur injoignable")
		{
			$IpColor = "'#B3B3B3'"
			$StdColor = "'#B3B3B3'"
		}
		else
		{
			$IpColor = "'#99CEEE'"
			$StdColor = "'#99CEEE'"
		}
		Add-Content $Report "<td bgcolor=$StdColor align=center><B>$($Obj.Name)</B></td>"
		Add-Content $Report "<td bgcolor=$IpColor align=center><B>$($Obj.ipv4Address)</B></td>"
		Add-Content $Report "<td bgcolor=$IpColor align=center><B>$($Obj.LastLogonDate)</B></td>"  
		Add-Content $Report "<td bgcolor=$StdColor align=center><B>$($Obj.Description)</B></td>" 	
		Add-Content $Report "<td bgcolor=$StdColor align=center><B>$($Obj.User)</B></td>" 	
		Add-Content $Report "<td bgcolor=$StdColor align=center><B>$($Obj.LastUserLogon)</B></td>" 					
		Add-Content $Report "</tr>"
	}
	
	
	# Finalization of the integrity report
	Add-Content $Report "</table>" 
	Add-Content $Report "</body>" 
	Add-Content $Report "</html>"
}
#############################################################################################


#############################################################################################
# Calling functions																			#
#############################################################################################
Create
Corps
ListWks
Clean
FilerCopy
#############################################################################################
