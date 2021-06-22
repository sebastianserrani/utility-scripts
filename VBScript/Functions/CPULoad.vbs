'===============================================================================
' Name: CPULoad
' Input:
'    None
' Output:
'    Int
' Purpose: Returns the CPU load
' Usage:
'     Wscript.Echo "CPU load: " & CPULoad
'===============================================================================

function CPULoad()
	Dim strComputer, objWMI, objCPU, Proc, intProcessorCount, intProcessorLoad
	strComputer = "." 'This computer
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objCPU = objWMI.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE name != '_Total'",,48)
	intProcessorCount = 0
	intProcessorLoad = 0
	For Each Proc In objCPU
		intProcessorLoad = intProcessorLoad + cdbl(Proc.PercentProcessorTime) + cdbl(proc.PercentPrivilegedTime)
		intProcessorCount = intProcessorCount + 1
	Next
	If intProcessorLoad > 0 And intProcessorCount > 0 Then
		intProcessorLoad = (intProcessorLoad / intProcessorCount)
	End If
	intProcessorLoad = Replace(intProcessorLoad, ",", ".")
	CPULoad = intProcessorLoad
End Function
