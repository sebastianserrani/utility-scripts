<#
.SYNOPSIS
    Check LOCAL users for upcoming password expiration.

.EXAMPLE
    powershell -ExecutionPolicy ByPass -File .\List-LocalUSerPasswordExpiration.ps1

.LINK
    To Disable Password Expiration for All User Accounts on PC (https://www.sevenforums.com/tutorials/73210-password-expiration-enable-disable.html)
    https://social.technet.microsoft.com/Forums/scriptcenter/en-US/b69912e9-c6c9-4de3-9e3e-8d4530095935/adsi-for-local-accounts
    https://gist.github.com/jdhitsolutions/fd12000ca43811674122
    https://stackoverflow.com/questions/51827241/password-expiring-email-notification-local-account-powershell-script
    https://berkenbile.wordpress.com/2013/04/26/manage-local-user-accounts-with-powershell/
    https://4sysops.com/archives/a-password-expiration-reminder-script-in-powershell/
    https://community.spiceworks.com/topic/1957028-scrip-to-show-days-remaining-fro-password-to-expire

.NOTES
    Works with powershell 2.0 too.
#>

clear
Write-Host "Check local user's password expiration"

$ComputerName = $Env:COMPUTERNAME
$now = Get-Date
$Days = 7 #days ahead the user will be notified of the expiry
$Count = 0; #count active users

Foreach($Computer in $ComputerName){
	$AllLocalAccounts = Get-WmiObject -Class Win32_UserAccount -Namespace "root\cimv2" `
    -Filter "LocalAccount='$True' AND Disabled='$False'" -ComputerName $Computer -ErrorAction Stop
	Foreach($LocalAccount in $AllLocalAccounts){
		$user = ([adsi]"WinNT://$computer/$($LocalAccount.Name),user")
		$username = [string]$user.Name
		$fullname = [string]$user.FullName
		$pwAge    = $user.PasswordAge.Value
		$maxPwAge = $user.MaxPasswordAge.Value
		$pwLastSet = $now.AddSeconds(-$pwAge)
		$remainingsecs = $maxPwAge - $pwAge
		$expirydate = $now.AddSeconds($remainingsecs)
		$expirydatelong = ($now.AddSeconds($remainingsecs)).ToLongDateString()
		$DaysToExpire = (New-TimeSpan -Start $now -End $expirydate).Days
		$Flags = $user.UserFlags.psbase.Value
		$result = ""		
		If (-Not ($Flags -band 65536)){ #Password does expire
			$Count += 1
			$AgeDays = $user.PasswordAge.psbase.Value / 86400
			$MaxAge = $user.MaxPasswordAge.psbase.Value / 86400
			If ($AgeDays -gt $MaxAge){
				$result = "Expired"
				Write-Host "[$Count] User: $username, remainig: $DaysToExpire, expires: $expirydatelong, menssage: $result" -ForegroundColor Red
			} ElseIf (($AgeDays + $Days) -gt $MaxAge) {
				$result = "Expires on $DaysToExpire days"
				Write-Host "[$Count] User: $username, remainig: $DaysToExpire, expires: $expirydatelong, menssage: $result" -ForegroundColor Yellow
			} Else {
				$result = ""
				Write-Host "[$Count] User: $username, remainig: $DaysToExpire, expires: $expirydatelong" -ForegroundColor Green
			}
		}
	}
}

Write-Host "Done."
