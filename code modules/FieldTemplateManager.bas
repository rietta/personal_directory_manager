Attribute VB_Name = "FieldTemplateManager"
Option Explicit
Public CurrentTemplateFile As String
Public CurrentTemplateFolder As String
Public TemplateIndex() As String
Dim TempBackupName As String
Public TemplateData(100, 12) As String

' Load Template Index from File
' Created on 3-12-1999
Public Sub LoadTemplateIndex()
    Dim i As Integer, NumberOfTemplates As Integer
    NumberOfTemplates = MyIni_GetInt(CurrentTemplateFile, "INDEX", "NumberOfTemplates")
    If NumberOfTemplates >= 1 Then
        ReDim TemplateIndex(NumberOfTemplates - 1) As String
        For i = 1 To NumberOfTemplates
            TemplateIndex(i - 1) = GetMyIni(CurrentTemplateFile, "INDEX", Format$(i))
        Next i
    End If
End Sub
' Initialize Information
' Created on 3-12-1999
Public Sub InitializeTemplateInfo()
    ' The default template file is called
    ' templates.pdt and is located in tge
    ' same folder as the PDM exe file.
    CurrentTemplateFolder = App.Path
    
    If right(CurrentTemplateFolder, 1) <> "\" Then CurrentTemplateFolder = CurrentTemplateFolder & "\"
    CurrentTemplateFile = CurrentTemplateFolder & "templates.pdt"
    TempBackupName = CurrentTemplateFolder & "backup of " & GetFName(CurrentTemplateFile)
    
    If Dir$(CurrentTemplateFile) = "" And Dir$(TempBackupName) <> "" Then
        MsgBox "The template file could not be found, but I was able to recover some template information from the backup file.", vbInformation, SPROGRAMNAME
        FileCopy TempBackupName, CurrentTemplateFile
    ElseIf Dir$(CurrentTemplateFile) = "" Then
        MsgBox "The template file, templates.pdt, could not be found in this folder", vbInformation, SPROGRAMNAME
        ReDim TemplateIndex(0) As String
    End If
    LoadTemplateIndex
End Sub

Public Sub LoadAllTemplateData(TempData() As String)
    Dim i As Integer, j As Integer
    
    ' Load all data from file
    ' The first array element (0) is reserved for the local file data inside this binders header
    For i = 0 To UBound(TemplateIndex)
        For j = 1 To 13
            TempData(i + 1, j - 1) = GetMyIni(CurrentTemplateFile, TemplateIndex(i), "Field" & Format(j))
        Next j
    Next i
End Sub

Public Sub SaveTemplateData(TempIndex() As String, TempData() As String)
    Dim i As Integer, NumTemp As Integer, j As Integer
    
    
    ' Rename Current File to a backup filename
    If Dir(TempBackupName) <> "" Then Kill TempBackupName
    Name CurrentTemplateFile As TempBackupName
    
    ' Create a New Template File
    i = FreeFile
    Open CurrentTemplateFile For Output As i
        Print #i, "' Personal Directory Manager Template File"
        Print #i, "' Copyright 1999 Rietta.  All Rights Reserved."
    Close i
    
    NumTemp = UBound(TempIndex) + 1
    SetMyIni CurrentTemplateFile, "INDEX", "NumberOfTemplates", Format(NumTemp)
    For i = 1 To NumTemp
            SetMyIni CurrentTemplateFile, "INDEX", Format(i), TempIndex(i - 1)
            For j = 1 To 13
                SetMyIni CurrentTemplateFile, TempIndex(i - 1), "Field" & Format(j), TempData(i, j - 1)
            Next j
        
    Next i
End Sub

Public Sub SetTemplateFieldsToActive(FieldTemplateIndex As Integer)
    Dim i As Integer
    LoadAllTemplateData TemplateData()
    For i = 0 To 12
        UserField(i) = TemplateData(FieldTemplateIndex + 1, i)
    Next i
End Sub
