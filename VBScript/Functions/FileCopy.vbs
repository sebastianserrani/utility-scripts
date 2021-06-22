'===============================================================================
' Name: FileCopy
' Input:
'    ByVal strSourceFilePath String - Source path of the file to copy
'    ByVal strDestPath String - Destination path, without the file name
'    ByVal booOverWrite Boolean - If it should overwrite a existing file
' Output:
'    None
' Purpose: Copy a file
' Author: Sebastian Serrani
'===============================================================================

Sub FileCopy(strSourceFilePath, strDestPath, booOverWrite)
	Dim strNewDestPath, strFileName, booOver
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	strNewDestPath = objFSO.BuildPath(objFSO.GetAbsolutePathName(strDestPath), "\")
    strFileName = objFSO.GetFileName(strSourceFilePath)
	booOver = CBool(booOverWrite)
	If Not objFSO.FileExists(strSourceFilePath) Then 'source file doesn't exists
		MsgBox("Source file doesn't exists: " & strSourceFilePath)
	ElseIf objFSO.FileExists(strNewDestPath & strFileName) And Not booOver Then 'destination file exists and shouldn't be overwritten
		MsgBox("Destination file already exists: " & strNewDestPath & strFileName)
	Else
		Err.clear
		On Error Resume Next
		objFSO.CopyFile strSourceFilePath, strNewDestPath & strFileName, booOver
		If Err.Number = 53 Then 'file not found
			MsgBox("Can't copy " & strFileName & " to " & strNewDestPath & ": [" & Err.Number & "] " & Err.Description)
		ElseIf Err.Number <> 0 Then 'error copying
			MsgBox("Can't copy  " & strFileName & " to " & strNewDestPath & ": [" & Err.Number & "] " & Err.Description)
		Else
			MsgBox("File copied: " & strFileName)
		End If
		On Error GoTo 0
	End If 
End Sub
