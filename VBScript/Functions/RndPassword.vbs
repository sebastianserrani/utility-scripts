'===============================================================================
' Name: RndPassword
' Input:
'    ByVal intLength Int - Length of the password
' Output:
'    String
' Purpose: Generates a semi random password
' Remarks: WARNING: it's not safe to generate a password with this!
' Usage:
'     Set objShell = WScript.CreateObject("WScript.Shell")
'     intLogType = ErrorType(Clng(Err.Number))
'     objShell.LogEvent intLogType, "Result [ " & Err.Number & " ] " & Err.Description
'     Err.clear
'===============================================================================

Function RndPassword(intLength)
	Dim intMax, k, intValue, strChar, strName
	' Specify the alphabet of characters to use.
	Const Chars = "abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ!#$%&()=@:;[]*?="
	Randomize()
    strName = ""
    For k = 1 To intLength
        ' Retrieve random digit
        intValue = Fix(71 * Rnd())
        ' Convert to character in allowed alphabet.
        strChar = Mid(Chars, intValue + 1, 1)
        ' Build the name.
        strName = strName & strChar
    Next
	RndPassword = strName
End Function
