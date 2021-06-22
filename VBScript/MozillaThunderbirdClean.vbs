'===============================================================================
' Name: MozillaThunderbirdClean
' Input:
'    None
' Output:
'    None
' Purpose: Script to delete *.mozmsgs folders, which contain *.mozeml and *.wdseml files.
'          They are used for Windows Search integration.
' Remarks: Can be disable by unchecking Tools > Options > Advanced > General > 
'          > System Integration > Allow Windows Search to search messages
'          http://kb.mozillazine.org/Files_and_folders_in_the_profile_-_Thunderbird
' Author: Sebastian Serrani
' Modified: 2019-05-02
'===============================================================================

ForceCScriptExecution
Set stdIn = WScript.stdin
Set stdOut = WScript.stdout
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strJobFolder : strJobFolder = objShell.expandEnvironmentStrings("%APPDATA%") & "\Thunderbird\Profiles"
Dim booConfirm, objPrompResult

strJobFolder = InputBox("Enter the path to Thunderbird profiles:", "MozillaThunderbirdClean", strJobFolder)
objPrompResult = MsgBox("Do you want to delete the global index (global-messages-db.sqlite) and the .mozmsgs folders found in the Thunderbird profiles path? If you answer NO, the folders will only be shown, without deleting anything.", vbYesNoCancel + vbQuestion + vbDefaultButton1, "MozillaThunderbirdClean")
Select Case objPrompResult
	Case vbYes
		booConfirm = True
	Case vbNo
		booConfirm = False
	Case vbCancel
		WScript.Quit
End Select

If booConfirm = True Then
	Wscript.Echo "ATTENTION! The global index (global-messages-db.sqlite) and the found .mozmsgs folders will be deleted!"
Else
	Wscript.Echo "No file will be deleted!"
End If

Wscript.Echo
Wscript.Echo "Searching in directory: " & strJobFolder
FindSubfolder(strJobFolder)
WScript.StdIn.ReadLine

Sub FindSubfolder(strPath)
	Err.clear
    Set objFolder = objFSO.GetFolder(strPath)
	On Error Resume Next
	If  Not objFolder.SubFolders Is Nothing And _
		Not IsNull(objFolder.SubFolders.Count) And _
		Not IsEmpty(objFolder.SubFolders.Count) Then	
		For Each objFile In objFSO.GetFolder(objFolder.Path).Files
			stdOut.write "."
			If objFile.Name = "global-messages-db.sqlite" Then
				stdOut.writeblanklines 1				
				If booConfirm = True Then
					Wscript.Echo "Database deleted: " & objFile.Path
					objFSO.DeleteFile(objFile.Path)
				Else
					Wscript.Echo "Database found: " & objFile.Path
				End If				
			End if
		Next
		If objFolder.SubFolders.Count > 0 Then
			For Each objSubFolder In objFolder.SubFolders
				stdOut.write "."
				If InStr(objSubFolder.Name, "mozmsgs") > 0 Then
					stdOut.writeblanklines 1					
					If booConfirm = True Then
						Wscript.Echo "Folder deleted: " & objSubFolder.Path
						objFSO.DeleteFolder(objSubFolder.Path)
					Else
						Wscript.Echo "Folder found: " & objSubFolder.Path
					End If
					
				End If
				Call FindSubfolder(objSubFolder.Path)
			Next
		End If
	End If
	If Err.Number <> 0 Then
		stdOut.writeblanklines 1
		Wscript.Echo "Error looking for subfolders in: " & objFolder.Path
	End If
	On Error Goto 0 
End Sub

Sub ForceCScriptExecution()
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run _
            "cscript //nologo """ & _
            WScript.ScriptFullName & _
            """ " & Str
        WScript.Quit
    End If
End Sub