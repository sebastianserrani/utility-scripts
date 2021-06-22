'===============================================================================
' Name: MemoryFreePhysical
' Input:
'    None
' Output:
'    Lng
' Purpose: Returns the total free memory of the computer in MB
' Author: Sebastian Serrani
' Usage:
'     Wscript.Echo "Free physical memory: " & FormatNumber(MemoryFreePhysical,0,,,0)
'===============================================================================

function MemoryFreePhysical()
	Dim strComputer, lngResult, objWMI, colSettings, ObjOS
	strComputer = "." 'This computer
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colSettings = objWMI.ExecQuery("Select * from Win32_OperatingSystem")
	For Each ObjOS in colSettings 
	    lngResult = ObjOS.FreePhysicalMemory
	Next
	MemoryFreePhysical = lngResult/1024
End function
