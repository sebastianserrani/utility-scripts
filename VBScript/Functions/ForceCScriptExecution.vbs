'===============================================================================
' Name: ForceCScriptExecution
' Input:
'    None
' Output:
'    None
' Purpose: Forces the script to run in a command-line environment
' Author: Mark Unwin <marku@opmantek.com> in Open-AudIT
'===============================================================================

Sub ForceCScriptExecution()
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run _
            "cscript //nologo """ & _
            WScript.ScriptFullName & _
            """ " & Str
        WScript.Quit
    End If
End Sub