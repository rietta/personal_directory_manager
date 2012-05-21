Attribute VB_Name = "Module5"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Program Launcher 1.0
' Copyright 1998 ER Software
' All Rights Reserved
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' THIS CODE IS AN INTERNAL TOOL AND SHOULD NEVER BE
' DISTRIBUTED IN THIS UNCOMPILED FORM.

Option Explicit

Type progtype
    Title As String
    location As String
    version As String
End Type

Global programs() As progtype

' Returns:
' True if refrence was deleted
' False if it wasn't deleted
' -2 if it wasn't found.
'
' Use FindRefrence to find the refrence number of a program
Function DeleteRefrence(pi As Integer) As Integer

End Function

' Returns:
' The refrence index of the program which titled matched prog
' If no match can be found it returns -2
Function FindRefrence(prog As String) As Integer
    Dim i As Integer
    prog = Trim$(LCase$(prog))
    
    On Error Resume Next
    For i = 0 To UBound(programs)
        If prog = Trim$(LCase$(programs(i).Title)) Then
             FindRefrence = i
             Exit Function
        End If
    Next i
    FindRefrence = -2
End Function

Sub GetProgramInfo(ByVal retStr As String, prog As progtype)

 ' Parse the string information
 Dim l As Integer
 
 l = InStr(retStr, ";")
 prog.Title = Left$(retStr, l - 1)
 retStr = right$(retStr, Len(retStr) - l)

 l = InStr(retStr, ";")
 prog.version = Left$(retStr, l - 1)
 retStr = right$(retStr, Len(retStr) - l)
 
 l = InStr(retStr, ";")
 prog.location = Left$(retStr, l - 1)

 prog.Title = Trim$(prog.Title)
 prog.version = Trim$(prog.version)
 prog.location = Trim$(prog.location)

End Sub

' Returns the index of the first unused registry place.
' If all registries are used then it returns -1.
Function GetUnusedRefrence() As Integer
    Dim i As Integer, Fetch As String
    For i = 0 To 30
        Fetch = GetMyIni("ERSOFT.INI", "Programs", Trim$(Str$(i)))
        Fetch = Trim$(Fetch)
        If Fetch = Null Or Fetch = "" Then
                GetUnusedRefrence = i
                Exit Function
        End If
    Next i
    GetUnusedRefrence = -1
End Function

Sub LaunchProgram(pi As Integer)
 ' Start the program in array index PI
 On Error Resume Next
 Dim X As Integer, File As String
 File = programs(pi).location
 ChDir GetPath(File)
 X = Shell(File, 1)
 If Err Then MsgBox "Lauch was unable to start " & programs(pi).Title, MB_ICONEXCLAMATION, "ER Launcher"
End Sub

'
'  Loads all registered programs into an array.
'  progName is the name of the current program
'
Function LoadPrograms(progName As String)
 Dim temp As String, i As Integer, loaded As Integer
 Dim theprog As progtype
 loaded = 0
 For i = 0 To 30
     temp = GetMyIni("ERSOFT.INI", "Programs", Trim$(Str$(i)))
     temp = Trim$(temp)
     If temp <> "" Then
           GetProgramInfo temp, theprog
           If LCase$(theprog.Title) <> progName Then
                ReDim Preserve programs(loaded) As progtype
                programs(loaded).Title = theprog.Title
                programs(loaded).version = theprog.version
                programs(loaded).location = theprog.location
                loaded = loaded + 1
           End If
     End If
 Next i
 LoadPrograms = loaded
End Function

