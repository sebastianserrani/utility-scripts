<#
.SYNOPSIS
    Checks if users have expired passwords.

.EXAMPLE
    powershell -ExecutionPolicy ByPass -File .\Reset-ActiveDirectoryUserPasswordExpiration.ps1

.NOTES
    Works with powershell 2.0 too.
#>

$MaxDaysToExpire = 30 # The accounts that expire today or in the next X days will be reset.

function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) {
	#https://www.bgreco.net/powershell/format-color/
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	foreach($line in $lines) {
		$color = ''
		foreach($pattern in $Colors.Keys){
			if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if($color) {
			Write-Host -ForegroundColor $color $line
		} else {
			Write-Host $line
		}
	}
}

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

clear
Write-Host "Active Directory user passwords expiration check"

$GetAdUserParams = @{ 
    'Filter' = { (Enabled -eq $True) -and (PasswordNeverExpires -eq $false) } 
    'Properties' = 'Displayname', 'PasswordLastSet', 'PasswordExpired', 'PasswordNeverExpires','EmailAddress', 'msDS-UserPasswordExpiryTimeComputed'
} 
$GetAdUserParams.SearchBase = "OU=Users,DC=lan,DC=company,DC=com"

Get-ADUser @GetAdUserParams | Select-Object -Property @{Expression={($_.samAccountName).substring(($_.samAccountName).length - 3,3).ToUpper()};Label='OU';}, @{Expression={$_.samAccountName};Label='Account';}, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, @{Name="DaysToExpire";Expression={(New-TimeSpan -Start (get-date) -End ($_.PasswordLastSet + (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge)).Days}}, 'PasswordExpired' | Sort-Object -Property OU, samAccountName  | Format-Table -AutoSize -Wrap | Format-Color @{'\s([^\-0-9]\d{1}|1\d)\s' = 'Yellow'; '\s(-\d{1,}|0|True)\s' = 'Red' }


$confirmation = Read-Host "Reset the expiration of all accounts that expire in the next $($MaxDaysToExpire) days? (y/n)"
if ($confirmation -eq 's' -or $confirmation -eq 'y') {
    
	$ExpiredUsers = Get-ADUser @GetAdUserParams | Where {((NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed'))).Days -lt $MaxDaysToExpire) -or ($_.PasswordExpired -eq $true)}
	
	Write-Host "`n"
	foreach ($User in $ExpiredUsers) {
		Try {
			Set-ADUser -Identity "$($User.samAccountName)" -Replace @{pwdLastSet="0"}
			Set-ADUser -Identity "$($User.samAccountName)" -Replace @{pwdLastSet="-1"}
			Write-Host "Reset: $($User.samAccountName)`r"
		} Catch {
			Write-Host "Error reseting $($User.samAccountName): $($_.Exception.Message)`r" -ForegroundColor Red
		}
	}

	Write-Host "`r`nProcess finished.`r`n" -ForegroundColor Cyan
}
Write-Host "`n"
