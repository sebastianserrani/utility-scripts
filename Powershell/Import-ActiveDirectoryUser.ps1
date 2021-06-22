<#
.SYNOPSIS
    Imports Active Directory User accounts from a CSV file.

.EXAMPLE
    powershell -ExecutionPolicy ByPass -File .\Import-ActiveDirectoryUser.ps1

.LINK
    Based on https://www.toddklindt.com/blog/Lists/Posts/Post.aspx?ID=362

.NOTES
    CSV format expected:
    firstname,lastname,username,createuser,enabled,email,otheremail,streetaddress,city,zipcode,state,country,department,password,passwchange,telephone,jobtitle,company,ou
    Jhon,Doe,jdoe,1,1,jdoe@company.com,jdoe-alt@company.com,1000 Street Ave,San Francisco,94158,CA,US,Department,PassW0rd!,0,555-5555,Title,Company,"OU=Employees,OU=Users,OU=Company,DC=lan,DC=company,DC=com"
#>


Import-Module ActiveDirectory -ErrorAction SilentlyContinue

$dnsroot = "@company.com"

$users = Import-Csv AD-User-import.csv

Write-Output "Import users to Active Directory"

foreach ($user in $users) {
	$username = $user.username
	if (Get-ADUser -F {SamAccountName -eq $username}) {
		Write-Warning "A user account with username $username already exist."
	} else {
		if ($user.createuser -ne 1){
			Write-Host "A user account with username $username was SKIPED." -ForegroundColor Yellow
		} else {
			try {
				$params = @{
					SamAccountName = $username
					UserPrincipalName = ($username + $dnsroot)
					Name = ($user.firstname + " " + $user.lastname)
					GivenName = $user.firstname
					Surname = $user.lastname
					Enabled = [System.Convert]::ToBoolean([int]($user.enabled))
					DisplayName = ($user.firstname + " " + $user.lastname)
					Path = $user.ou
					Company = $user.company
					State = $user.state
					Country = $user.country
					AccountPassword = (ConvertTo-SecureString ($user.password) -AsPlainText -force)
					ChangePasswordAtLogon = [System.Convert]::ToBoolean([int]($user.passwchange))
				}
                if ($user.email -ne ""){
					$params.Add("EmailAddress", $user.email)
				}                
				if ($user.otheremail -ne ""){
					$params.Add("OtherAttributes", @{"otherMailbox" = $user.otheremail; "info" = $user.otheremail})
				}
				if ($user.city -ne ""){
					$params.Add("City", $user.city)
				}
				if ($user.streetaddress -ne ""){
					$params.Add("StreetAddress", $user.streetaddress)
				}
				if ($user.zipcode -ne ""){
					$params.Add("PostalCode", $user.zipcode)
				}
				if ($user.telephone -ne ""){
					$params.Add("OfficePhone", $user.telephone)
				}
				if ($user.department -ne ""){
					$params.Add("Department", $user.department)
				}
				if ($user.jobtitle -ne ""){
					$params.Add("Title", $user.jobtitle)
				}
				New-ADUser @params #-verbose
				Write-Host "A user account with username $username was added." -ForegroundColor Green
			}
			catch [System.Object]
			{
				Write-Warning "Could not create user $username, $_"
				Write-Host ($params | Out-String) -ForegroundColor Yellow
				Write-Warning "`n"
			}
		}
	}
}

Write-Output "Finished importing users to Active Directory."
