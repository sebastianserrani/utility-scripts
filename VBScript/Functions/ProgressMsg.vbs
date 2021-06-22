'===============================================================================
' Name: ProgressMsg
' Input:
'    ByVal strMessage String - Message to display
'    ByVal strButtons String - Message box buttons
'    ByVal strWindowTitle String - Window title
' Output:
'    None
' Purpose: Displays a non-blocking progress message box that the originating script can kill
'          If StrMessage is blank, take down previous progress message box
'          Using 4096 in Msgbox below makes the progress message float on top of things
' Remarks: must have Dim objProgressMsg at the top of thescript for this to work as described
' Author: Denis St-Pierre
' Usage:
'     ProgressMsg "Program X update found. Copying new files...", "+vbExclamation", "Update"
'     ProgressMsg "", "", "Update"
'     ProgressMsg "Update complete.", "+vbInformation", "Update"
'===============================================================================

Function ProgressMsg(strMessage, strButtons, strWindowTitle)
    Dim strTEMP, strTempVBS, objTempMessage
	strTEMP = objShell.ExpandEnvironmentStrings( "%TEMP%" )
    If strMessage = "" Then
        ' Disable Error Checking in case objProgressMsg doesn't exists yet
        On Error Resume Next
        ' Kill ProgressMsg
        objProgressMsg.Terminate( )
        ' Re-enable Error Checking
        On Error Goto 0
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs" 'Control File for reboot

    ' Create Message.vbs, True=overwrite
    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "MsgBox """ & strMessage & """, 4096" & strButtons & ", """ & strWindowTitle & """" )
    objTempMessage.Close

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0

    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = objShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )

    'Set wshShell = Nothing
    'Set objFSO   = Nothing
End Function