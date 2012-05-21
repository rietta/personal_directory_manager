VERSION 5.00
Begin VB.Form frmNotePad 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Untitled"
   ClientHeight    =   3540
   ClientLeft      =   1590
   ClientTop       =   2370
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "NOTEPAD.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   5835
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFPrinterSetup 
         Caption         =   "Printer &Setup"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFSB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit Notebook"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuESb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuECut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuECopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuESelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuETime 
         Caption         =   "Time/&Date"
      End
   End
   Begin VB.Menu mnuCharacter 
      Caption         =   "&Text"
      Begin VB.Menu mnuCRegular 
         Caption         =   "Regular"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCSB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFontOptions 
         Caption         =   "Bold"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCFontOptions 
         Caption         =   "Italic"
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuCFontOptions 
         Caption         =   "Underline"
         Index           =   2
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuCSB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOFontS 
         Caption         =   "&Fonts..."
      End
      Begin VB.Menu mnuOFontSize 
         Caption         =   "Font &Size"
         Begin VB.Menu mnuFontSize 
            Caption         =   "8"
            Index           =   0
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "9"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "10"
            Index           =   2
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "11"
            Index           =   3
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "12"
            Index           =   4
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "13"
            Index           =   5
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "14"
            Index           =   6
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "16"
            Index           =   7
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "17"
            Index           =   8
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "18"
            Index           =   9
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "19"
            Index           =   10
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "20"
            Index           =   11
         End
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuSFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuOStatusbar 
         Caption         =   "&Statusbar"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu MnuWTileV 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuH 
         Caption         =   "Help Topics"
         Index           =   0
      End
      Begin VB.Menu mnuH 
         Caption         =   "Search For Help on..."
         Index           =   1
      End
      Begin VB.Menu mnuH 
         Caption         =   "How to Use Help"
         Index           =   2
      End
      Begin VB.Menu mnuH 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuH 
         Caption         =   "About..."
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFontSize As Integer

Private Sub cmdCancel_Click()

End Sub

Private Sub Form_Activate()
ShowDocumentStatus
End Sub

Private Sub Form_Load()
 CurrentFontSize = 1
 Fetch$ = GetMyIni("BDMAN.INI", "Notebook", "Font")
 Fetch$ = Trim$(Fetch$)
 If Fetch$ <> "" Then
    SFont$ = Fetch$
    Text1.FontName = SFont$
 End If
 Text1.ForeColor = QBColor(0)
 GetRecentFiles
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Msg, filename, NL
    Dim Response As Integer

    Me.WindowState = 0
    
    
    If FState(Me.Tag).Dirty Then
        filename = Me.Caption
        NL = Chr$(10) & Chr$(13)
        Msg = "The text in [" & filename & "] has changed."
        Msg = Msg & NL
        Msg = Msg & "Do you want to save the changes?"
        Response = MsgBox(Msg, 51, frmMDI.Caption)
        Select Case Response
        ' User selects Yes
        Case 6
           'Call the Save procedure
           Call mnuFSave_Click
           If FState(Me.Tag).Dirty = True Then Cancel = True 'User Cancled Save
        ' User selects No
        ' Ok to unload
        Case 7
            Cancel = False
        ' User selects Cancel
        ' Cancel the unload
        Case 2
            Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 And ScaleHeight <> 0 Then
        Text1.Visible = False
        Text1.Height = ScaleHeight
        Text1.Width = ScaleWidth
        Text1.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FState(Me.Tag).Deleted = True
    
'Hide toolbar edit buttons if no notepad windows
    If Not AnyPadsLeft() Then
      ChangeToolState False
      frmMDI.lblFont.Caption = ""
    End If
End Sub

Private Sub mnuCFontOptions_Click(Index As Integer)
mnuCFontOptions(Index).Checked = Not mnuCFontOptions(Index).Checked
Text1.FontBold = mnuCFontOptions(0).Checked
Text1.FontItalic = mnuCFontOptions(1).Checked
Text1.FontUnderline = mnuCFontOptions(2).Checked
End Sub

Private Sub mnuCharacter_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuCRegular_Click()
 For i% = 0 To 2
   mnuCFontOptions(i%).Checked = False
 Next i%
Text1.FontBold = False
Text1.FontItalic = False
Text1.FontUnderline = False

End Sub

Private Sub mnuECopy_Click()
    EditCopyProc
End Sub

Private Sub mnuECut_Click()
    EditCutProc
End Sub

Private Sub mnuEDelete_Click()
  ' If cursor is not at the end of the notepad.
  If Screen.ActiveControl.SelStart <> Len(Screen.ActiveControl.text) Then
    ' If nothing is selected, extend selection by one.
    If Screen.ActiveControl.SelLength = 0 Then
      Screen.ActiveControl.SelLength = 1
      ' If cursor is on a blank line, extend selection by two.
      If Asc(Screen.ActiveControl.SelText) = 13 Then
        Screen.ActiveControl.SelLength = 2
      End If
    End If
    ' Delete selected text.
    Screen.ActiveControl.SelText = ""
  End If
End Sub

Private Sub mnuEdit_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuEPaste_Click()
    EditPasteProc
End Sub

Private Sub mnuESelectAll_Click()
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.text)
End Sub

Private Sub mnuETime_Click()
    frmFormats.Show 1
    If SelFormat.SFormat <> "" Then Text1.SelText = SelFormat.SFormat
End Sub

Private Sub mnuEUndo_Click()
  Dim UndoResult
  UndoResult = SendMessage(Text1.hwnd, EM_UNDO, 0, 0)
End Sub

Private Sub mnuFClose_Click()
    Unload Me
End Sub

Private Sub mnufexit_Click()
    ' Unloading the MDI form invokes the QueryUnload event
    ' for each child form, then the MDI form - before unloading
    ' the MDI form. Setting the Cancel argument to True in any of the
    ' QueryUnload events aborts the unload.

    Unload frmMDI


End Sub

Private Sub mnuFile_Click()
    If Text1.text = "" Then
        mnuFPrint.Enabled = False
    Else
        mnuFPrint.Enabled = True
    End If
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuFNew_Click()
    SelectTemplate
End Sub

Private Sub mnuFontSize_Click(Index As Integer)
 mnuFontSize(CurrentFontSize).Checked = False
 CurrentFontSize = Index
 mnuFontSize(CurrentFontSize).Checked = True
 s$ = mnuFontSize(CurrentFontSize).Caption
 Text1.FontSize = Val(Trim$(s$))
End Sub

Private Sub mnuFOpen_Click()
    FOpenProc
End Sub

Private Sub mnuFPrint_Click()
    If Text1.text <> "" Then
        PrintDocument Me
    Else
        MsgBox "There is no text to print.", MB_ICONEXCLAMATION, "Notebook"
    End If
End Sub

Private Sub mnuFPrinterSetup_Click()
x% = BDM_PrinterSetup()
End Sub

Private Sub mnuFSave_Click()
 SaveDocument Me
End Sub

Private Sub mnuFSaveAs_Click()
    Dim SaveFileName As String

    SaveFileName = GetFileName()
    If SaveFileName <> "" Then SaveFileAs (SaveFileName)
    ' Update the recent files menu
    UpdateFileMenu (SaveFileName)
End Sub

Private Sub mnuH_Click(Index As Integer)
   HelpMenu Index
End Sub

Private Sub mnuHelp_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuOFonts_Click()
 nPrintColor = FState(Me.Tag).color
 Fonts.Show 1
 If SFont = "" Then Exit Sub
 If SetAsDefault Then SetMyIni "BDMAN.INI", "Notebook", "Font", SFont$
 SetAsDefault = False
 Text1.FontName = SFont
 frmMDI.lblFont.Caption = SFont
 frmMDI.lblFont.ForeColor = QBColor(nPrintColor)
 Text1.ForeColor = QBColor(nPrintColor)
 FState(Me.Tag).color = nPrintColor
End Sub

Private Sub mnuOptions_Click()
    frmMDI.lblStatus = "Ready"
    mnuOToolbar.Checked = frmMDI!picToolbar.Visible
End Sub

Private Sub mnuOStatusbar_Click()
     OptionsToolbarProc Me, False
End Sub

Private Sub mnuOToolbar_Click()
    OptionsToolbarProc Me, True
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
  OpenFile (mnuRecentFile(Index).Caption)
  ' Update recent files list for new notepad.
  GetRecentFiles
End Sub

Private Sub mnuSearch_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuSFind_Click()
    If Me!Text1.SelText <> "" Then
        frmFind!Text1.text = Me!Text1.SelText
    Else
        frmFind!Text1.text = findstring
    End If
    gFirstTime = True
    frmFind.Show 1
End Sub

Private Sub mnuSFindNext_Click()
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSFind_Click
    End If
End Sub

Private Sub mnuWArrange_Click()
    frmMDI.Arrange ARRANGE_ICONS
End Sub

Private Sub mnuWCascade_Click()
    frmMDI.Arrange CASCADE
End Sub

Private Sub mnuWindow_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuWTileH_Click()
frmMDI.Arrange TILE_HORIZONTAL
End Sub

Private Sub MnuWTileV_Click()
frmMDI.Arrange TILE_VERTICAL
End Sub

Private Sub Text1_Change()
    FState(Me.Tag).Dirty = True
    ShowDocumentStatus
End Sub

Private Sub Text1_GotFocus()
    If frmFind.Visible Then
        frmFind.ZOrder 0
    End If
    If Me.Tag <> "" Then
       CurrentDocument = Me.Tag
       Font = Text1.FontName
       frmMDI.lblFont.Caption = Font
       nPrintColor = FState(Me.Tag).color
       frmMDI.lblFont.ForeColor = QBColor(nPrintColor)
       frmMDI.lblFont.FontName = Text1.FontName
    Else
       CurrentDocument = 0
    End If
End Sub

Private Sub Text1_MouseMove(button As Integer, Shift As Integer, x As Single, Y As Single)
ShowDocumentStatus
End Sub

