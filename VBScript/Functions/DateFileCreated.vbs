'===============================================================================
' Name: DateFileCreated
' Input:
'    ByVal strFilePath String
' Output:
'    Date
' Purpose: Returns the creation date of a file
' Author: Sebastian Serrani
'===============================================================================

Function DateFileCreated(strFilePath)
   Dim objFSO, objFile, datFileCreated
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(strFilePath)
   DateFileCreated = objFile.DateCreated
End Function
