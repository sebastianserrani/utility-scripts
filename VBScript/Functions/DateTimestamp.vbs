'===============================================================================
' Name: DateTimestamp
' Input:
'    ByVal datDate Date
' Output:
'    Int
' Purpose: Returns the seconds from Unix Epoch on January 1st, 1970 to datDate
' Author: Sebastian Serrani
'===============================================================================

Function DateTimestamp(datDate)
	DateTimestamp = DateDiff("s", "01/01/1970 00:00:00", datDate)
End Function
