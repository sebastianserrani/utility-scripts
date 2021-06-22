'===============================================================================
' Name: TimeSpan
' Input:
'    ByVal dt1 String - Initial date time
'    ByVal dt2 String - Final date time
' Output:
'    Str
' Purpose: Returns the time span in hours:minutes:seconds format
' Usage:
'     Dim datToday : datToday = Date()
'     Dim StartTime : StartTime = "00:00:01"
'     Dim EndTime : EndTime = "23:59:59"
'     Dim datStartDatetime : datStartDatetime = "11/01/2021 " & StartTime
'     Dim datEndDatetime : datEndDatetime = datToday & " " & EndTime
'     Wscript.Echo TimeSpan(datStartDatetime, datEndDatetime)
'===============================================================================

Function TimeSpan(dt1, dt2)
	Dim seconds, minutes, hours
    If (isDate(dt1) And IsDate(dt2)) = false Then
        TimeSpan = "00:00:00"
        Exit Function
    End If

    seconds = Abs(DateDiff("S", dt1, dt2))
    minutes = seconds \ 60
    hours = minutes \ 60
    minutes = minutes mod 60
    seconds = seconds mod 60
    if len(hours) = 1 then hours = "0" & hours

    TimeSpan = hours & ":" & _
        RIGHT("00" & minutes, 2) & ":" & _
        RIGHT("00" & seconds, 2)
End Function
