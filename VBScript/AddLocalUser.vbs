'===============================================================================
' Name: AddLocalUser
' Input:
'    ByVal strUser String - User name
'    ByVal strPass String - Password
'    ByVal strAdminGroupName String - User group name
'    ByVal booPasswordExpires Bool - Password expires?
' Output:
'    None
' Purpose: Adds a new local user
' Remarks: WARNING: password manipulation without any security whatsoever!
' Author: Sebastian Serrani
' Usage:
'     AddLocalUser("JohnDoe", "MyPass123456", "Administrators", False)
'===============================================================================

Sub AddLocalUser(ByVal strUser, ByVal strPass, ByVal strGroupName, ByVal booPasswordExpires)
	Dim objShell, objEnv, strComputer, colAccounts, objUser
	Dim objPasswordExpirationFlag
	Set objShell = CreateObject("Wscript.Shell")

	Set objEnv = objShell.Environment("Process")
	strComputer = objEnv("COMPUTERNAME")

	Set colAccounts = GetObject("WinNT://" & strComputer & ",computer")

	Set objUser = colAccounts.Create("user", strUser)
	objUser.SetPassword strPass
	
	If booPasswordExpires = False Then
		Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
		objPasswordExpirationFlag = ADS_UF_DONT_EXPIRE_PASSWD
		objUser.Put "userFlags", objPasswordExpirationFlag
	End If
	
	objUser.SetInfo 
	Set Group = GetObject("WinNT://" & strComputer & "/" & strGroupName & ",group")
	Group.Add(objUser.ADspath)
End Sub
