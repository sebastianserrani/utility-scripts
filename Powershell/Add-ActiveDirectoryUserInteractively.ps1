<#
.SYNOPSIS
    Interactively create an Active Directory User account

.EXAMPLE
    powershell -ExecutionPolicy ByPass -File .\Add-ActiveDirectoryUserInteractively.ps1

.LINK
    Based on https://www.toddklindt.com/blog/Lists/Posts/Post.aspx?ID=362
#>

Import-Module ActiveDirectory

clear
Write-Host "Interactive user creation in Active Directory"

$dnsroot = "@company.com"

$username = Read-Host 'User name? (jdoe)'
if (Get-ADUser -F {SamAccountName -eq $username}) {
	Write-Host "This user already exists. Exiting..." -ForegroundColor Yellow
} else {
	$email = Read-Host 'Email? (jdoe@company.com)'
	$firstname = Read-Host 'First name? (Jhon)'
	$lastname = Read-Host 'Last name? (Doe) (Optional)'
	$password = Read-Host 'Password? (Min. 8 characters)' -AsSecureString
	$company = Read-Host 'Company? (Company Name)'
	$state = Read-Host 'State? (New York)'
	$country = Read-Host 'Country? (US)'
	if (
		($email -ne "") -and
		($firstname -ne "") -and
		($password -ne "") -and
		($company -ne "") -and
		($state -ne "") -and
		($country -ne "")			
	){
		Read-Host -Prompt "Press any key to continue or CTRL + C to exit."		
		try {
			$params = @{
				SamAccountName = $username
				UserPrincipalName = ($username + $dnsroot)
				GivenName = $firstname
				Enabled = $true
				Path = "OU=Users,DC=lan,DC=CompanyName,DC=com"
				Company = $company
				State = $state
				Country = $country
				AccountPassword = $password
				ChangePasswordAtLogon = $false
				EmailAddress = $email
			}
			if ($lastname -ne ""){
				$params.Add("Surname", $lastname)
			}
			if (($firstname -ne "") -and ($lastname -ne "")){
				$params.Add("Name", $firstname + " " + $lastname)
				$params.Add("DisplayName", $firstname + " " + $lastname)
			} else {
				$params.Add("Name", $firstname + $lastname)
				$params.Add("DisplayName", $firstname + $lastname)
			}
			New-ADUser @params #-verbose
			Write-Host "User $username was created." -ForegroundColor Green
		}
		catch [System.Object]
		{
			Write-Warning "Could not create user $username, $_"
			Write-Host ($params | Out-String) -ForegroundColor Yellow
			Write-Warning "`n"
		}
	} else {
		Write-Host "Incomplete data. Cannot create user. Exiting..." -ForegroundColor Yellow
	}		
}
Write-Output "Done."
