Attribute VB_Name = "CHTMLHelp"
' This module is designed to interface with the new
' HTML based help system.  This program does not use
' the new Microsoft HTML Help with is dependent on
' Internet Explorer, but rather, loads help in the
' user's default web browser.

Option Explicit

Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Const SW_SHOWNORMAL = 1

Public HTMLHelpFolder As String
Sub OpenURL(Url As String)
    Dim x
    x = ShellExecute(0&, vbNullString, Url, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Public Function GetHTMLHelpFolder()
    Dim htmlpath As String
    htmlpath = App.Path
    If right$(htmlpath, 1) <> "\" Then htmlpath = htmlpath & "\"
    htmlpath = htmlpath & "help\"
    GetHTMLHelpFolder = htmlpath
End Function

Public Sub HTMLHelp(File As String)
    OpenURL GetHTMLHelpFolder() & File
End Sub

Public Sub NoHelp()
    HTMLHelp "nohelp.htm"
End Sub
