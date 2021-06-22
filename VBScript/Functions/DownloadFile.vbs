'===============================================================================
' Name: DownloadFile
' Input:
'    ByVal strFileURL Int - HTTP URL of the file
'    ByVal strDestFile Int - Path of destination
'    ByVal booShowErrors Bool - Show errors?
' Output:
'    Bool
' Purpose: Downloads a file
' Usage:
'     If Download("https://web.url/file.jpg", "C:\Temp\file.jpg", True) = True Then
'         WScript.Echo "File downloaded OK."
'     End If
'===============================================================================

Function DownloadFile(strFileURL, strDestFile, booShowErrors)
	Err.clear
	On Error Resume Next	
	Dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
	Dim bStrm: Set bStrm = createobject("Adodb.Stream")
	xHttp.Open "GET", strFileURL, False
	xHttp.Send
	With bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile strDestFile, 2 '//overwrite
	End With
	Set xHttp = Nothing
	Set bStrm = Nothing
	If Err.Number <> 0 Then
		If booShowErrors = True Then
			MsgBox  "Error downloading the file " & strFileURL & ":" & vbCrLf & _
					"[" & Err.Number & "] " & Err.Description
		End If
		Err.clear
		DownloadFile = False
	Else
		DownloadFile = True
	End If 
End Function
