'===============================================================================
' Name: DateFileLastMod
' Input:
'    ByVal strFilePath String
' Output:
'    Date
' Purpose: Returns the last modification date of a file
' Author: Sebastian Serrani
'===============================================================================

Function DateFileLastMod(strFilePath)
   Dim objFSO, objFile, datFileLastMod
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(strFilePath)
   DateFileLastMod = objFile.DateLastModified
End Function
