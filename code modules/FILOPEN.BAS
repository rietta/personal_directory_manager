Attribute VB_Name = "NotebookFileIOManager"

Sub FOpenProc()
    Screen.MousePointer = 11
    Dim retVal
    On Error Resume Next
    Dim OpenFileName As String
    frmMDI.CMDialog1.Filter = "Text Files(*.TXT)|*.txt|Batch Files(*.BAT)|*.bat|INI Files(*.INI)|*.INI|All Files(*.*)|*.*;*."
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.FileName = ""
    frmMDI.CMDialog1.Action = 1
    If Err <> 32755 Then 'user pressed cancel
        OpenFileName = frmMDI.CMDialog1.FileName
        OpenFile (OpenFileName)
    End If
    Screen.MousePointer = 0
End Sub

Function GetFileName()
    'Displays a Save As dialog and returns a file name
    'or an empty string if the user cancels
    Screen.MousePointer = 11
    On Error Resume Next
    frmMDI.CMDialog1.Filter = "Text Files(*.TXT)|*.txt|Batch Files(*.BAT)|*.bat|INI Files(*.INI)|*.INI|All Files(*.*)|*.*;*."
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.FileName = ""
    frmMDI.CMDialog1.FLAGS = OFN_OVERWRITEPROMPT
    frmMDI.CMDialog1.DefaultExt = "txt"
    frmMDI.CMDialog1.Action = 2
    If Err <> 32755 Then      'User cancelled dialog
        GetFileName = frmMDI.CMDialog1.FileName
    Else
        GetFileName = ""
    End If
    Screen.MousePointer = 0
End Function

Function OnRecentFilesList(FileName) As Integer
  Dim i

  For i = 1 To 4
    If frmMDI.mnuRecentFile(i).Caption = FileName Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function

Sub OpenFile(FileName)
    Dim NL, TextIn, GetLine, fn As Integer
    Dim fIndex As Integer
    frmMDI.lblStatus = "Starting new document..."
    NL = Chr$(13) + Chr$(10)
    
    On Error Resume Next
    fn = FreeFile
    ' open the selected file
    Open FileName For Input As #fn
    If Err Then
        MsgBox "Can't open file: " + FileName, , SPROGRAMNAME
        Exit Sub
    End If
    ' change mousepointer to an hourglass
    frmMDI.lblStatus = "Opening document..."
    Screen.MousePointer = 11
    ' change form's caption and display new text
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = GetFName(UCase$(FileName))
    FState(fIndex).filepath = UCase$(FileName)
    Document(fIndex).Show
    Document(fIndex).Text1.text = Input$(LOF(fn), fn)
    FState(fIndex).Dirty = False
    Close #fn
    ChangeToolState True
    ' reset mouse pointer
    UpdateFileMenu UCase$(FileName)
    Screen.MousePointer = 0
    frmMDI.lblStatus = ""
End Sub

Sub SaveDocument(Doc As Form)
    Dim FileName As String

    If Left(Doc.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet,
        ' get the filename, then call the
        ' save procedure
        FileName = GetFileName()
    Else
        ' The caption contains the name of the open file
        FileName = FState(Val(Doc.Tag)).filepath
        FileName = Trim(FileName)
    End If
    ' call the save procedure, if Filename = Empty then
    ' the user selected Cancel in the Save As dialog, otherwise
    ' save the file
    Screen.MousePointer = 11
    If FileName <> "" Then SaveFileAs FileName
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(FileName)
On Error Resume Next
    Dim Contents As String, fn As Integer
    frmMDI.lblStatus = "Saving " & FileName & "..."
    frmMDI.lblStatus.Refresh
    fn = FreeFile
    ' open the file
    Open FileName For Output As #fn
    ' put contents of the notepad into a variable
    Contents = frmMDI.ActiveForm.Text1.text
    ' display hourglass
    Screen.MousePointer = 11
    ' write variable contents to saved file
    Print #fn%, Contents
    Close #fn%
    ' reset the mousepointer
    Screen.MousePointer = 0
    ' set the Notepad's caption

    If Err Then
        MsgBox Error, 48, SPROGRAMNAME
    Else
        frmMDI.ActiveForm.Caption = GetFName(UCase$(FileName))
        ' reset the dirty flag
        FState(frmMDI.ActiveForm.Tag).Dirty = False
        FState(frmMDI.ActiveForm.Tag).filepath = FileName
    End If
    frmMDI.lblStatus = "Ready"
End Sub

Sub UpdateFileMenu(FileName)
        Dim retVal
        ' Check if OpenFileName is already on MRU list.
        retVal = OnRecentFilesList(FileName)
        If Not retVal Then
          ' Write OpenFileName to MDINOTEPAD.INI
          WriteRecentFiles (FileName)
        End If
        ' Update menus for most recent file list.
        GetRecentFiles
End Sub

