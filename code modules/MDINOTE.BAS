Attribute VB_Name = "NotebookManager"
Option Explicit
Type SelFormatType
    SFormat As String
    SIndex As Integer
End Type
Global SelFormat As SelFormatType

Function AnyPadsLeft() As Integer
    Dim i As Integer

    ' Cycle throught the document array.
    ' Return True if there is at least one
    ' open document remaining.
    For i = 1 To UBound(Document)
        If Not FState(i).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next

End Function

Sub CenterForm(frmParent As Form, frmChild As Form)
' This procedure centers a child form over a parent form.
' Calling this routine loads the dialog. Use the Show method
' to display the dialog after calling this routine ( ie MyFrm.Show 1)

Dim l, t
  ' get left offset
  l = frmParent.Left + ((frmParent.Width - frmChild.Width) / 2)
  If (l + frmChild.Width > Screen.Width) Then
    l = Screen.Width = frmChild.Width
  End If

  ' get top offset
  t = frmParent.Top + ((frmParent.Height - frmChild.Height) / 2)
  If (t + frmChild.Height > Screen.Height) Then
    t = Screen.Height - frmChild.Height
  End If

  ' center the child formfv
  frmChild.Move l, t

End Sub

Sub ChangeToolState(St As Integer)
    frmMDI!imgCutButton.Visible = St
    frmMDI!imgCopyButton.Visible = St
    frmMDI!imgPasteButton.Visible = St
    frmMDI!imgSaveButton.Visible = St
    frmMDI!imgPrintButton.Visible = St
    If St Then
       frmMDI!imgMainWinButton.Left = 2880
    Else
       frmMDI!imgMainWinButton.Left = 840
    End If
End Sub

Sub EditCopyProc()
    ' Copy selected text to Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
End Sub

Sub EditCutProc()
    ' Copy selected text to Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Delete selected text.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc()
    ' Place text from Clipboard into active control.
    frmMDI.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub

Sub FileNew()
    Dim fIndex As Integer
    frmMDI.lblStatus = "Starting new document..."
    Screen.MousePointer = 11
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Untitled:" & fIndex
    Document(fIndex).Show
    FState(fIndex).filepath = ""
    FState(fIndex).Dirty = False
    ' Make sure toolbar edit buttons are visible
    ChangeToolState True
    DoEvents
    ShowDocumentStatus
    Screen.MousePointer = 0
End Sub

Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle throught the document array. If one of the
    ' documents has been deleted, then return that
    ' index.
    For i = 1 To ArrayCount
        If FState(i).Deleted Then
            FindFreeIndex = i
            FState(i).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.

    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt()
    Dim Start, pos, findstring, sourcestring, msg, Response, offset
    
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        offset = 1
    Else
        offset = 0
    End If

    If gFirstTime Then offset = 0

    Start = frmMDI.ActiveForm.ActiveControl.SelStart + offset
        
    If gFindCase Then
        findstring = gFindString
        sourcestring = frmMDI.ActiveForm.ActiveControl.text
    Else
        findstring = UCase(gFindString)
        sourcestring = UCase(frmMDI.ActiveForm.ActiveControl.text)
    End If
            
    If gFindDirection = 1 Then
        pos = InStr(Start + 1, sourcestring, findstring)
    Else
        For pos = Start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If

    ' If string is found
    If pos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = pos - 1
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(findstring)
    Else
        msg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        Response = MsgBox(msg, 0, SPROGRAMNAME)
    End If
    
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False

End Sub

Sub GetRecentFiles()
  Dim retVal, key As String, i, j
  Dim IniString As String

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from BDMAN.INI
  For i = 1 To 4
    key = "RecentFile" & i
    retVal = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "BDMAN.INI")
    If retVal And Left(IniString, 8) <> "Not Used" Then
      ' Update the MDI form's menu.
      frmMDI.mnuRecentFile(0).Visible = True
      frmMDI.mnuRecentFile(i).Caption = IniString
      frmMDI.mnuRecentFile(i).Visible = True
  
      ' Iterate through all the notepads and update each menu.
      For j = 1 To UBound(Document)
        If Not FState(j).Deleted Then
          Document(j).mnuRecentFile(0).Visible = True
          Document(j).mnuRecentFile(i).Caption = IniString
          Document(j).mnuRecentFile(i).Visible = True
        End If
      Next j
    End If
  Next i

End Sub

Sub HelpMenu(i As Integer)
 Select Case i
   Case 0: 'ShowHelpContents
   Case 1: 'SearchHelp
   Case 2: 'ShowHelpOnHelp
   Case 3: 'Seperator bar
   Case 4: AboutBox.Show 1
 End Select

End Sub

Sub NewMemo()
 If NoteBook = 0 Then frmMDI.Show
 If frmMDI.WindowState = 1 Then frmMDI.WindowState = 0 'Maximize WIndow to normal size
 FileNew
 SetCRLF
 frmMDI.ActiveForm.Text1.text = "Date: " & Format$(Date$, "mmmm d, yyyy") & crlf & "To: " & crlf & "From: " & crlf & "Subject: " & crlf
 GetRecentFiles
End Sub

Sub OptionsToolbarProc(CurrentForm As Form, TS As Integer)
    If TS Then
       CurrentForm.mnuOToolbar.Checked = Not CurrentForm.mnuOToolbar.Checked
    Else
       CurrentForm.mnuOStatusbar.Checked = Not CurrentForm.mnuOStatusbar.Checked
    End If
    
    If TypeOf CurrentForm Is MDIForm Then
    Else
        frmMDI.mnuOToolbar.Checked = CurrentForm.mnuOToolbar.Checked
    End If
    
    If CurrentForm.mnuOToolbar.Checked Then
        frmMDI.picToolbar.Visible = True
        SetMyIni "BDMAN.INI", "Notebook", "Toolbar", Str$(True)
    Else
        frmMDI.picToolbar.Visible = False
        SetMyIni "BDMAN.INI", "Notebook", "Toolbar", Str$(False)
    End If

    If CurrentForm.mnuOStatusbar.Checked Then
        frmMDI.Picture1.Visible = True
        SetMyIni "BDMAN.INI", "Notebook", "Statusbar", Str$(True)
    Else
        frmMDI.Picture1.Visible = False
        SetMyIni "BDMAN.INI", "Notebook", "Statusbar", Str$(False)
    End If
End Sub

Sub PrintDocument(Doc As Form)
 Dim Notes As String, i As Integer, fn As String
 Screen.MousePointer = 11
 frmPrintDocument.Show 1
 If Not DoPrint Then Exit Sub
 
 fn$ = Doc.Caption
 frmMDI.lblStatus = "Printing " & fn$ & "..."
 frmMDI.lblStatus.Refresh
'-----------------------------------------------------------------
'Process and print line
'-----------------------------------------------------------------
 Screen.MousePointer = 11
 On Error Resume Next
 Printer.ForeColor = Doc.Text1.ForeColor
 Printer.FontBold = Doc.Text1.FontBold
 Printer.FontItalic = Doc.Text1.FontItalic
 Printer.FontUnderline = Doc.Text1.FontUnderline
 Printer.FontName = Doc.Text1.FontName
 Printer.FontSize = Doc.Text1.FontSize
 
 For i% = 1 To 4
   Printer.Print    'Add Top Margin
 Next i%
 Notes$ = Doc.Text1.text & crlf
 PrintLine Notes$ 'Use Wordwrap
 Printer.EndDoc
 Screen.MousePointer = 0
 frmMDI.lblStatus = "Ready"
End Sub

Sub SelectTemplate()
   ' Creates a new document the current template the user
   ' selects

   Dim msg As String
   Load frmFormats
   frmFormats.Caption = "Document Templates"
   frmFormats.Label1.Caption = "Availible Templates:"
   frmFormats.lstFormats.Clear
   frmFormats.lstFormats.AddItem "Normal"
   frmFormats.lstFormats.AddItem "Memo"
   frmFormats.lstFormats.AddItem "Personal Letter"
   frmFormats.lstFormats.AddItem "Business Letter"
   frmFormats.lstFormats.AddItem "Letter Wizard"
   frmFormats.lstFormats.ListIndex = 0
   frmFormats.Show 1
   If SelFormat.SFormat = "" Then Exit Sub   'User chose cancel

   Select Case SelFormat.SIndex
       Case 0: FileNew  ' Create new window
       Case 1
           FileNew   ' Create new window
           frmMDI.ActiveForm.Text1.text = "Date: " & Format$(Date$, "mmmm d, yyyy") & crlf & "To: " & crlf & "From: " & crlf & "Subject: " & crlf
       Case 2
           FileNew   ' Create new window
           msg = Format$(Date$, "dddd, mmmm d, yyyy") & crlf & crlf
           msg = msg & "[Place the persons address here]" & crlf & crlf
           msg = msg & "Dear [Addresses name here]," & crlf & crlf
           msg = msg & "[Place your body here]" & crlf & crlf
           msg = msg & "Sincerely,"
           frmMDI.ActiveForm.Text1.text = msg
       Case 3
           FileNew   ' Create new window
           msg = Format$(Date$, "dddd, mmmm d, yyyy") & crlf & crlf
           msg = msg & "[Place the persons address here]" & crlf & crlf
           msg = msg & "Dear [Addresses name here]:" & crlf & crlf
           msg = msg & "[Place your body here]" & crlf & crlf
           msg = msg & "Sincerely,"
           frmMDI.ActiveForm.Text1.text = msg
       Case 4
           LetterWizard.Show 1
           If SelFormat.SFormat <> "" Then
               FileNew   ' Create new window
               frmMDI.ActiveForm.Text1.text = SelFormat.SFormat
           End If
   End Select
End Sub

Sub ShowDocumentStatus()
Dim temp As String
temp = LCase$(FState(frmMDI.ActiveForm.Tag).filepath)
If temp = "" Then temp = frmMDI.ActiveForm.Caption
If Not FState(frmMDI.ActiveForm.Tag).Dirty Then temp = temp & " - unmodified"
If frmMDI.lblStatus <> temp Then frmMDI.lblStatus = temp
End Sub

Sub WriteRecentFiles(OpenFileName As String)
  Dim i, j, key As String, retVal
  Dim IniString As String
  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 3 To 1 Step -1
    key = "RecentFile" & i
    retVal = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "BDMAN.INI")
    If retVal And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFile" & (i + 1)
      retVal = WritePrivateProfileString("Recent Files", key, IniString, "BDMAN.INI")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retVal = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, "BDMAN.INI")

End Sub

