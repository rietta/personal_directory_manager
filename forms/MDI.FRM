VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Notebook"
   ClientHeight    =   5730
   ClientLeft      =   75
   ClientTop       =   1050
   ClientWidth     =   9255
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   270
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   9225
      TabIndex        =   1
      Top             =   5460
      Width           =   9255
      Begin VB.Label lblFont 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   6840
         TabIndex        =   3
         Top             =   0
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   8640
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327680
      End
      Begin VB.Image imgMainWinButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         Picture         =   "MDI.frx":030A
         Tag             =   "Main Window - Show the main window"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgPrintButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Picture         =   "MDI.frx":0470
         Tag             =   "Print - Print the current document - Ctrl+P"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgSaveButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         Picture         =   "MDI.frx":05D6
         Tag             =   "Save - Save the current document - Ctrl+S"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgPasteButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Picture         =   "MDI.frx":073C
         Tag             =   "Paste - insert the clipboards contents at the insertion point - Ctrl+V"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgCopyButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         Picture         =   "MDI.frx":08A2
         Tag             =   "Copy - copy the selected text to the clipboard - Ctrl+C"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgCutButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         Picture         =   "MDI.frx":0A08
         Tag             =   "Cut - move or cut selected text to the clipboard - Ctrl+X"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgFileOpenButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         Picture         =   "MDI.frx":0B6E
         Tag             =   "Open - Retrieve an existing document into a new window - Ctrl+O"
         Top             =   30
         Width           =   285
      End
      Begin VB.Image imgFileNewButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   30
         Picture         =   "MDI.frx":0CD4
         Tag             =   "New Document - Create a new document in a new window - Ctrl+N"
         Top             =   30
         Width           =   285
      End
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
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuOStatusbar 
         Caption         =   "&Statusbar"
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
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HelpTips As Integer
Dim HelpTipOpen As Integer
Dim CurrentCaption As String
Dim bttn As Double, bttnPush As Integer

Private Sub imgCopyButton_Click()
    EraseHighlightedButton 1, bttn
    EditCopyProc
End Sub

Private Sub imgCopyButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgCopyButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgCopyButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 2.2 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 2.2
    End If
    Controlin3DonPicbox imgCopyButton, 1, 0, bttnPush, picToolbar
    ShowTip imgCopyButton
End Sub

Private Sub imgCopyButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgCopyButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgCutButton_Click()
    EraseHighlightedButton 1, bttn
    EditCutProc
End Sub

Private Sub imgCutButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgCutButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgCutButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 2.1 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 2.1
    End If
    Controlin3DonPicbox imgCutButton, 1, 0, bttnPush, picToolbar
    ShowTip imgCutButton
End Sub

Private Sub imgCutButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgCutButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgFileNewButton_Click()
    EraseHighlightedButton 1, bttn
    FileNew
End Sub

Private Sub imgFileNewButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgFileNewButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileNewButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.1 Then
        EraseHighlightedButton 1, bttn
        bttn = 1.1
    End If
    Controlin3DonPicbox imgFileNewButton, 1, 0, bttnPush, picToolbar
    ShowTip imgFileNewButton
End Sub

Private Sub imgFileNewButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgFileNewButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileOpenButton_Click()
    EraseHighlightedButton 1, bttn
    FOpenProc
End Sub

Private Sub imgFileOpenButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgFileOpenButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileOpenButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.2 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 1.2
    End If
    Controlin3DonPicbox imgFileOpenButton, 1, 0, bttnPush, picToolbar
    ShowTip imgFileOpenButton
End Sub

Private Sub imgFileOpenButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgFileOpenButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgMainWinButton_Click()
    EraseHighlightedButton 1, bttn
    Screen.MousePointer = 11
    If frmMain.WindowState = 1 Then frmMain.WindowState = 0
    frmMain.Show
    NoteBook = 1
    If ViewWindowOpen Then ViewWindow.ZOrder (0) ' Place View Window in front of other windows
    Screen.MousePointer = 0
End Sub

Private Sub imgMainWinButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgMainWinButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgMainWinButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 3.1 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 3.1
    End If
    Controlin3DonPicbox imgMainWinButton, 1, 0, bttnPush, picToolbar
    
    ShowTip imgMainWinButton
End Sub

Private Sub imgMainWinButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgMainWinButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgPasteButton_Click()
    EraseHighlightedButton 1, bttn
    EditPasteProc
End Sub

Private Sub imgPasteButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgPasteButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgPasteButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 2.3 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 2.3
    End If
    Controlin3DonPicbox imgPasteButton, 1, 0, bttnPush, picToolbar
    ShowTip imgPasteButton
End Sub

Private Sub imgPasteButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgPasteButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgPrintButton_Click()
    If Document(CurrentDocument).WindowState <> 1 And Document(CurrentDocument).Text1.text <> "" Then
        EraseHighlightedButton 1, bttn
        PrintDocument frmMDI.ActiveForm 'Print the current document
    Else
        MsgBox "There is no text to print.", MB_ICONEXCLAMATION, "Notebook"
    End If
End Sub

Private Sub imgPrintButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgPrintButton_MouseMove button, Shift, x, y
    If Document(CurrentDocument).WindowState = 1 Or Document(CurrentDocument).Text1.text = "" Then Exit Sub
End Sub

Private Sub imgPrintButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.4 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 1.4
    End If
    Controlin3DonPicbox imgPrintButton, 1, 0, bttnPush, picToolbar
    ShowTip imgPrintButton
End Sub

Private Sub imgPrintButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgPrintButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub imgSaveButton_Click()
    EraseHighlightedButton 1, bttn
    SaveDocument frmMDI.ActiveForm
End Sub

Private Sub imgSaveButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgSaveButton_MouseMove button, Shift, x, y
    
    If Document(CurrentDocument).WindowState = 1 Then Exit Sub
    frmMDI.lblStatus = "Save Current Document..."
End Sub

Private Sub imgSaveButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.3 Then
        EraseHighlightedButton 1, bttn
        bttnPush = False
        bttn = 1.3
    End If
    Controlin3DonPicbox imgSaveButton, 1, 0, bttnPush, picToolbar
    ShowTip imgSaveButton
End Sub

Private Sub imgSaveButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgSaveButton_MouseMove button, Shift, x, y
    frmMDI.lblStatus = "Ready"
End Sub

Private Sub MDIForm_Load()
    Dim Fetch As String
    NoteResize = True
    Screen.MousePointer = 11
    Fetch = GetMyIni("BDMAN.INI", "Notebook", "Toolbar")
    If Fetch <> "" Then mnuOToolbar.Checked = Not Val(Trim(Fetch))
    Fetch = GetMyIni("BDMAN.INI", "Notebook", "Statusbar")
    If Fetch <> "" Then mnuOStatusbar.Checked = Not Val(Trim(Fetch))
    OptionsToolbarProc Me, True
    OptionsToolbarProc Me, False
    
    NoteBook = 1    'Notebook is Running
    lblStatus = "Starting new document..."
    ' Application starts here (Load event of Startup form).
    Show
    ' Always set working directory to directory containing the application.
    ChDir App.Path
    
    'Initialize document form arrays, and show first document.
    ReDim Document(1)
    ReDim FState(1)
    Document(1).Tag = 1
    FState(1).Dirty = False
    Document(1).Show

    ' Read BDMAN.INI and set recent file menu items appropriately
    GetRecentFiles
    Screen.MousePointer = 0
    lblStatus = ""
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If NoteBook = 2 Then
        End 'User exited the main program
    End If
    NoteBook = 0 'Notebook is not running
End Sub

Private Sub mnufexit_Click()
    Unload frmMDI
End Sub

Private Sub mnuFile_Click()
frmMDI.lblStatus = "Ready"
End Sub

Private Sub mnuFNew_Click()
    SelectTemplate
End Sub

Private Sub mnuFOpen_Click()
    FOpenProc
End Sub

Private Sub mnuH_Click(Index As Integer)
    HelpMenu Index
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
  ' Update recent files list.
  GetRecentFiles
End Sub

Private Sub picToolbar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 0 Then
        EraseHighlightedButton 1, bttn
        bttn = 0
    End If

    frmMDI.lblStatus = "Ready"
End Sub

Private Sub Picture1_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
frmMDI.lblStatus = "Ready"
End Sub

Private Sub ShowTip(Con As Control)
 Tip$ = Con.Tag
 Cap$ = Me.Caption
 If Trim$(Cap$) <> Tip$ Then frmMDI!lblStatus.Caption = Tip$
 HelpTipOpen = True
End Sub

