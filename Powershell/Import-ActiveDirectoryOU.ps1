<#
.SYNOPSIS
    Import Active Directory OUs

.EXAMPLE
    powershell -ExecutionPolicy ByPass -File .\Import-ActiveDirectoryOU.ps1

#>

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

clear
Write-Output "Import OUs on Active Directory"

$Domain = "DC=lan,DC=CompanyName,DC=com"
$struct = "CompanyName
CompanyName\Branch1
CompanyName\Branch1\Users
CompanyName\Branch1\Users\Employees
CompanyName\Branch1\Users\Managers
CompanyName\Branch2
CompanyName\Branch2\Users
CompanyName\Branch2\Users\Employees
CompanyName\Branch2\Users\Managers
CompanyName\Externals
CompanyName\Externals\Users".Split("`r")

Foreach($Property In $struct){
	$PropertyValue = $Property.Trim()
	$Path = ""
	$OUs = (Split-Path $PropertyValue -Parent).Split("\")
	[array]::Reverse($OUs)
	$OUs | Foreach-Object {
		if ($_.Length -eq 0) {
			return
		}
		$Path = $Path + "OU=" + $_ + ","
	}
	$Path += $Domain
	$NewOU = Split-Path $PropertyValue -Leaf
	if (Get-ADOrganizationalUnit -Filter "distinguishedName -eq 'OU=$NewOU,$Path'") {
		Write-Warning "An OU with name $NewOU already exist (OU=$NewOU,$Path)."
	} else {
		try {
			New-ADOrganizationalUnit -Name "$NewOU" -Path "$Path" -ProtectedFromAccidentalDeletion $false
			Write-Host "Created a new OU $NewOU on path $Path" -ForegroundColor Green
		} catch [System.Object]	{
			Write-Warning "Could not create OU $NewOU, $_ `n"
		}
	}
}

Write-Host "Finished importing OUs."