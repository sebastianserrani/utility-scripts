'===============================================================================
' Name: DateTimeAmerican
' Input:
'    ByVal datDate Date
' Output:
'    String
' Purpose: Returns the datDate in american format (MM-DD-YYYY HH:MM:SS)
' Author: Sebastian Serrani
'===============================================================================

Function DateTimeAmerican(datDate)
    DateTimeAmerican = Right("0" & Month(datDate),2)  & "/" & _ 
	Right("0" & Day(datDate),2)  & "/" & _ 
	Year(datDate) & " " & _     
    Right("0" & Hour(datDate),2) & ":" & _
    Right("0" & Minute(datDate),2) & ":" & _  
	Right("0" & Second(datDate),2)
End Function
