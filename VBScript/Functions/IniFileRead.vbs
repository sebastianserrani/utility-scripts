'===============================================================================
' Name: IniFileRead
' Input:
'    ByVal strINIPath String - Path to the .ini file
' Output:
'    IniFileRead Dic - Dictionary with the config values
' Purpose: Reads an .ini files and returns a dictionary of values
' Usage:
'     Dim strWorkingDirectory : Set strWorkingDirectory = WScript.Arguments.Named("WorkingDirectory")
'     Dim dicConfig : Set dicConfig = IniFileRead(strWorkingDirectory & "\Config.ini")
'     If dicConfig.Count < 1 Then
'     	MsgBox "Error: The configuration file seems invalid."
'     	WScript.Quit
'     End If
'     Wscript.Echo "Config.ini > [SECTION] > Key: " & dicConfig("SECTION")("Key")
'===============================================================================

Function IniFileRead(strINIPath)
	Dim dicTmp : Set dicTmp = CreateObject("Scripting.Dictionary")
	Dim objIniFile : Set objIniFile = objFSO.OpenTextFile(strINIPath)
	Dim strLine, strSec, arrKV, strValue
	Do Until objIniFile.AtEndOfStream
		strLine = Trim(objIniFile.ReadLine())
		If "[" = Left(strLine, 1) Then
			strSec = Mid(strLine, 2, Len(strLine) - 2)
			Set dicTmp(strSec) = CreateObject("Scripting.Dictionary")
		Else
			If "" <> strLine And ";" <> Left(strLine, 1) And "/" <> Left(strLine, 1) Then
				arrKV = Split(strLine, "=")
				If 1 = UBound(arrKV) Then
					strValue = Trim(arrKV(1))
					If "'" = Left(strValue, 1) Or """" = Left(strValue, 1) Or "#" = Left(strValue, 1) Then
						dicTmp(strSec)(Trim(arrKV(0))) = Mid(strValue, 2, Len(strValue) - 2)
					Else
						dicTmp(strSec)(Trim(arrKV(0))) = strValue
					End If
				End If
			End If
		End If
	Loop
	objIniFile.Close
	Set IniFileRead = dicTmp
End Function
