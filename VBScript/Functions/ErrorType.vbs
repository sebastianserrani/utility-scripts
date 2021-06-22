'===============================================================================
' Name: ErrorType
' Input:
'    ByVal objErrorNumber Int - Err.Number input
' Output:
'    ErrorType int - Possible values:
'        0 = Success
'        1 = Error
'        2 = Warning
'        4 = Information
' Purpose: Returns the error type
' Remarks: Useful for logging events in Windows event logs
' Author: Sebastian Serrani
' Usage:
'     Set objShell = WScript.CreateObject("WScript.Shell")
'     intLogType = ErrorType(Clng(Err.Number))
'     objShell.LogEvent intLogType, "Result [ " & Err.Number & " ] " & Err.Description
'     Err.clear
'===============================================================================

Function ErrorType(ByVal objErrorNumber)
	Dim arrVBErrorCodes, booIsVBError, i
	booIsVBError = False
	arrVBErrorCodes = Array(5, 6, 7, 9, 10, 11, 13, 14, 28, 35, 48, 51, 53, 57, 58, 61, 67, 70, 75, 76, 91, 92, 94, 322, 424, 429, 430, 432, 438, 440, 445, 446, 447, 448, 449, 450, 451, 453, 455, 457, 458, 500, 501, 502, 503, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013, 1014, 1015, 1016, 1017, 1018, 1019, 1020, 1021, 1022, 1023, 1024, 1025, 1026, 1027, 1028, 1029, 1030, 1031, 1032, 1033, 1034, 1035, 1037, 1038, 1039, 1040, 1041, 1042, 1043, 1044, 1045, 1046, 32766, 32767, 32811)	
	For i = 0 to Ubound(arrVBErrorCodes)
		If arrVBErrorCodes(i) = objErrorNumber Then
			booIsVBError = True
			Exit For
		End If
	Next	
	Select Case True
		Case booIsVBError = True
			ErrorType = 1 'Error
		Case objErrorNumber <> 0
			ErrorType = 2 'Warning
		Case objErrorNumber = 0
			ErrorType = 0 'Success
		Case Else
			ErrorType = 4 'Information
	End Select
End Function
