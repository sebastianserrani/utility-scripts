'===============================================================================
' Name: CmdOutput
' Input:
'    ByVal strCommandLine String - Command to execute
' Output:
'    String - Command output or false
' Purpose: Returns the StdOut from a console command
' Author: Sebastian Serrani
' Usage:
'     'Capture idle sessions on a Remote Desktop Session Host server
'     arrIdleConnections = CmdOutput("qwinsta | find ""idle""")
'===============================================================================

Function CmdOutput(strCommandLine)
	Dim booOutput : booOutput = False 'Enable output from this function
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim strCommand, objTempFile, objReturnFile
	objTempFile = objFSO.GetTempName
	strCommand = "%COMSPEC% /c " & strCommandLine & " >" & objTempFile
	Err.Clear
	objShell.Run strCommand, 0, True
	If Err.Number <> 0 Then
		If booOutput = True Then
			Wscript.Echo "Error: " & Err.Number & " " & Err.Description
		End If
		CmdOutput = False
	Else
		If objFSO.FileExists(objTempFile) Then
			If objFSO.GetFile(objTempFile).Size > 0 Then
				Set objReturnFile = objFSO.OpenTextFile(objTempFile)
				CmdOutput = objReturnFile.Readall
				objReturnFile.Close
			Else
				CmdOutput = False
			End If
			objFSO.DeleteFile(objTempFile)
		Else
			CmdOutput = False
		End If
    End If
End Function
