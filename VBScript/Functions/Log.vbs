'===============================================================================
' Name: Log
' Input:
'    ByVal strText String - Text to log
' Output:
'    None
' Purpose: Logs to a .log file with the name of the script
' Author: Sebastian Serrani
' Usage:
'     Log "text!"
'     Log("text!")
'===============================================================================

Sub Log(ByVal strText)
	Dim booOutput : booOutput = False 'Enable output from this function
	Dim strLogFile, objLogFile, strDate 
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	strLogFile = objFSO.GetBaseName(Wscript.ScriptName) & ".log"
	strDate = Year(Now()) & "-" & _
	Right("0" & Month(Now()),2)  & "-" & _ 
	Right("0" & Day(Now()),2)  & " " & _	     
    Right("0" & Hour(Now()),2) & ":" & _
    Right("0" & Minute(Now()),2) & ":" & _  
	Right("0" & Second(Now()),2)
	On Error Resume Next
	Err.Clear
	Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True, -2)
	If Err.Number <> 0 Then
		If booOutput = True Then
			Wscript.Echo "Log: Error [" & Err.Number & "] " & Err.Description
		End If
	Else
		If booOutput = True Then
			Wscript.Echo "Log: " & strDate & " " & strText
		End If
		objLogFile.WriteLine(strDate & " " & strText)
		objLogFile.Close
	End If
End Sub
