'===============================================================================
' Name: ElevatePermissions
' Input:
'    None
' Output:
'    None
' Purpose: Relaunches the calling script asking for elevated permissions
' Author: Sebastian Serrani
' Modified: 2019-10-28
'===============================================================================

Sub ElevatePermissions()
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	strComputerUser = objShell.ExpandEnvironmentStrings("%USERNAME%")
	strWorkingDirectory = objFSO.GetParentFolderName(objFSO.GetFile(Wscript.ScriptFullName))
	
	If Not WScript.Arguments.Named.Exists("elevate") Then 'no admin
		CreateObject("Shell.Application").ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ /ComputerUser:" & strComputerUser & " /WorkingDirectory:" & strWorkingDirectory & " /elevate", "", "runas", 1
		WScript.Quit
	Else 'elevated
		Exit Sub
	End If
End Sub
