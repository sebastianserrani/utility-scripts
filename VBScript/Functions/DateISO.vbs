'===============================================================================
' Name: DateISO
' Input:
'    ByVal datDate Date
' Output:
'    String
' Purpose: Returns the Date in ISO 8601 format (YYYY-MM-DD)
' Author: Sebastian Serrani
'===============================================================================

Function DateISO(datDate)
	DateISO = Year(datDate) & "-" & _
	Right("0" & Month(datDate),2)  & "-" & _
	Right("0" & Day(datDate),2)
End Function
