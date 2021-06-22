'===============================================================================
' Name: MemoryTotal
' Input:
'    None
' Output:
'    Lng
' Purpose: Returns the total memory of the computer in MB
' Author: Sebastian Serrani
' Usage:
'     Wscript.Echo "Total memory: " & FormatNumber(MemoryTotal,0,,,0)
'===============================================================================

function MemoryTotal()
	Dim strComputer, lngResult, objWMI, colSettings, ObjOS
	strComputer = "." 'This computer
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colSettings = objWMI.ExecQuery("Select TotalPhysicalMemory from Win32_ComputerSystem")
	For Each ObjOS in colSettings 
	    lngResult = ObjOS.TotalPhysicalMemory 
	Next
	MemoryTotal = lngResult/(1024*1024) 'Memory in Mb
End function
