<#	
	.NOTES
	===========================================================================
	 Filename:     	Get_remove_old_profiles.ps1
	===========================================================================
	.DESCRIPTION
		Script to search for obsolete RDS profiles (user deleted or disabled in AD) prior to the year 2020;
        the script will then appropriate the rights to the folders and then delete them.
#>

Import-Module ActiveDirectory

$PF = "D:\PROFILS_RDS"
$ScriptFolder = "C:\Scripts\"
$Date = (Get-Date -Format yyyyMMdd-HHmmss)
$TempOld = "_OldProfile_PROFILS_RDS.csv"
$TempRem = "_ToRemove_PROFILS_RDS.csv"

$Global:TempReport = $ScriptFolder+$Date+$TempOld
$Global:RemoveReport = $ScriptFolder+$Date+$TempRem

$dir = dir $PF | select Name, LastWriteTime | sort LastWriteTime
$directory = $dir | Where-Object -Property LastWriteTime -NotMatch 2021 | Where-Object -Property LastWriteTime -NotMatch 2020 | sort Name

$Global:SAMList = ($directory.name).split(".") | Select-String -NotMatch "SELPRO" | Select-String -NotMatch "V2" | sort 
$SAMList | Out-File $Global:TempReport -Encoding utf8

foreach($Global:User in $Global:SAMList){
    
    $user = "$user"
    $test = [bool] (Get-ADUser -Filter {SamAccountName -eq $user} -ErrorAction SilentlyContinue)
    
    if($test -eq $False){
        ($user, $test) -join ";" | Out-File $Global:RemoveReport -Append -Encoding utf8}
    else {
        $get = Get-ADUser $user -Properties * -ErrorAction SilentlyContinue | Select Enabled
        $status = $get.Enabled
        ($user, $status) -join ";" | Out-File $Global:RemoveReport -Append -Encoding utf8
    }
}

$base = Import-Csv -Path $Global:RemoveReport -Delimiter ";" -Header "Name", "Enabled" -Encoding UTF8
$Global:basecheck = $base | Where-Object -Property "Enabled" -eq "False" | select Name

foreach($Global:profile in $Global:basecheck){

    $FName = $profile.name
    $endpath = ".selpro.v2"
    $Ufolder = $FName+$endpath
    $checkpath = gci $PF | Where-Object {$_.Name -eq $Ufolder} -ErrorAction SilentlyContinue | select FullName -ErrorAction SilentlyContinue
    
    if($checkpath){
    $testpath = Test-Path $checkpath.FullName -ErrorAction SilentlyContinue
        
        if($testpath){
        $part1 = $checkpath.FullName
        $path = "$part1\*"
        takeown.exe /F $part1 /a /r /d o
        icacls.exe $path /grant administrateurs:F /T /Q
        Remove-Item $part1 -Recurse -Force -Confirm:$False
        Write-Host -ForegroundColor Green "$part1 has been deleted"
    
        } else { 
        Write-Host -ForegroundColor Red "$FName > not deleted" }
        
    } else { 
    Write-Host -ForegroundColor Red "$FName > not deleted" }
}
