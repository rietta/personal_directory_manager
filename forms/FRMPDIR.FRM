VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "Personal Directory Manager"
   ClientHeight    =   5715
   ClientLeft      =   1065
   ClientTop       =   1155
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   3
   Icon            =   "FRMPDIR.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3960
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flexItems 
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   6240
      TabIndex        =   5
      Top             =   5460
      Width           =   6270
      Begin VB.Label lblStatus 
         Caption         =   "Copyright 1996-1999 Rietta."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.PictureBox picToolbar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   5895
      TabIndex        =   4
      Top             =   480
      Width           =   5955
      Begin VB.Image imgFindButton 
         Height          =   480
         Left            =   5040
         Picture         =   "FRMPDIR.frx":0442
         ToolTipText     =   "Search the binder for items matching a criteria"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblFindButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Find"
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   615
         Width           =   495
      End
      Begin VB.Image imgPrintButton 
         Height          =   315
         Left            =   4470
         Picture         =   "FRMPDIR.frx":0884
         ToolTipText     =   "Print the select item or all the items in this binder"
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblPrintButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Print"
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   615
         Width           =   615
      End
      Begin VB.Image imgViewButton 
         Height          =   480
         Left            =   3720
         Picture         =   "FRMPDIR.frx":0EAE
         ToolTipText     =   "View the selected item"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblViewButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&View"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   615
         Width           =   495
      End
      Begin VB.Label lblDeleteButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Delete"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   615
         Width           =   615
      End
      Begin VB.Image imgDeleteButton 
         Height          =   480
         Left            =   2880
         Picture         =   "FRMPDIR.frx":12F0
         ToolTipText     =   "Delete the selected item"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblEditButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   615
         Width           =   615
      End
      Begin VB.Image imgEditButton 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "FRMPDIR.frx":1732
         Picture         =   "FRMPDIR.frx":1A3C
         ToolTipText     =   "Edit the selected item"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblAddButton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Add"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imgAddButton 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "FRMPDIR.frx":1E7E
         Picture         =   "FRMPDIR.frx":2188
         ToolTipText     =   "Create a new Item"
         Top             =   100
         Width           =   480
      End
      Begin VB.Line lneTool1 
         X1              =   1440
         X2              =   1440
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Image imgFileNewButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   45
         Picture         =   "FRMPDIR.frx":25CA
         Tag             =   "New - Create a New Folder - Ctrl+N"
         ToolTipText     =   "New"
         Top             =   150
         Width           =   285
      End
      Begin VB.Image imgFileOpenButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         Picture         =   "FRMPDIR.frx":2730
         Tag             =   "Open - Open an Existing Folder - Ctrl+O"
         ToolTipText     =   "Open"
         Top             =   150
         Width           =   285
      End
      Begin VB.Image imgEditSortButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         Picture         =   "FRMPDIR.frx":2896
         Tag             =   "Sort - Sort Folder"
         ToolTipText     =   "Sort"
         Top             =   150
         Width           =   285
      End
      Begin VB.Image imgBookmarkButton 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Picture         =   "FRMPDIR.frx":29FC
         Tag             =   "View Bookmarks - Open Bookmark List"
         ToolTipText     =   "Bookmarks"
         Top             =   150
         Width           =   285
      End
   End
   Begin VB.PictureBox picListTools 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5895
      TabIndex        =   2
      Top             =   4800
      Width           =   5955
      Begin VB.Image imgRecycle 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5400
         Picture         =   "FRMPDIR.frx":2B62
         Tag             =   "Double-Click to open the Waste Bin."
         ToolTipText     =   "Waste-Bin"
         Top             =   40
         Width           =   480
      End
      Begin VB.Image imgNotebook 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4800
         Picture         =   "FRMPDIR.frx":2FA4
         Tag             =   "Double-Click to open the Notebook"
         ToolTipText     =   "Notebook"
         Top             =   40
         Width           =   480
      End
      Begin VB.Image imgPrintBuffer 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4080
         Picture         =   "FRMPDIR.frx":32AE
         Tag             =   "Double-Click to open the Print Buffer Manager"
         ToolTipText     =   "Print Buffer"
         Top             =   40
         Width           =   480
      End
      Begin VB.Image imgPlaceBookmark 
         Height          =   270
         Left            =   360
         Picture         =   "FRMPDIR.frx":35B8
         Tag             =   "Add a bookmark to the selected item."
         Top             =   20
         Width           =   285
      End
      Begin VB.Image imgListDown 
         Height          =   225
         Left            =   30
         Picture         =   "FRMPDIR.frx":3A32
         Tag             =   "Shift the Selected Item Down"
         ToolTipText     =   "Shift Selected Item Down"
         Top             =   255
         Width           =   225
      End
      Begin VB.Image imgListUp 
         Height          =   225
         Left            =   30
         Picture         =   "FRMPDIR.frx":3F66
         Tag             =   "Shift the Selected Item Up"
         ToolTipText     =   "Shift Selected Item Up"
         Top             =   5
         Width           =   225
      End
   End
   Begin VB.PictureBox picTitlebar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6270
      TabIndex        =   1
      Top             =   0
      Width           =   6270
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Line lneMenuBar 
         X1              =   -240
         X2              =   11880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image imgBinderHeader 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   120
         Picture         =   "FRMPDIR.frx":449A
         Tag             =   "Double-Click for more information on this binder."
         Top             =   10
         Width           =   480
      End
      Begin VB.Label lblFolder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<no binder is currently open>"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   0
         Tag             =   "Double-Click to change binder description."
         ToolTipText     =   "Shows the description of the binder.  Double-click to edit."
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Image imgRemoveBookmark 
      Height          =   270
      Left            =   6360
      Picture         =   "FRMPDIR.frx":48DC
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgAddBookmark 
      Height          =   270
      Left            =   6000
      Picture         =   "FRMPDIR.frx":4D56
      Tag             =   "Add a bookmark to the selected item."
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5280
      Picture         =   "FRMPDIR.frx":51D0
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4800
      Picture         =   "FRMPDIR.frx":5612
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   4
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close "
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFPSSB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Make Backup Copy"
      End
      Begin VB.Menu mnuExportSAT 
         Caption         =   "Export as &Text"
      End
      Begin VB.Menu mnuFMerge 
         Caption         =   "&Merge Binders"
      End
      Begin VB.Menu mnuSB1A 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFFolderPathInfo 
         Caption         =   "&Path Information..."
      End
      Begin VB.Menu FB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNotePad 
         Caption         =   "&Notebook"
         Begin VB.Menu mnuFileNotebook 
            Caption         =   "Open &Notebook"
            Index           =   0
         End
         Begin VB.Menu mnuFileNotebook 
            Caption         =   "New M&emo"
            Index           =   1
         End
         Begin VB.Menu mnuFileNotebook 
            Caption         =   "Open &Text File..."
            Index           =   2
         End
      End
      Begin VB.Menu FB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Printer &Setup"
      End
      Begin VB.Menu FB6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   5
      Begin VB.Menu mnuECopy 
         Caption         =   "&Copy Item"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEPaste 
         Caption         =   "&Paste Item"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuESendTo 
         Caption         =   "&Send To..."
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "&Print Buffer"
            Index           =   0
         End
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "&Clipboard"
            Index           =   1
         End
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "&Notebook"
            Index           =   2
         End
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "Send &E-Mail Address to Clipboard"
            Index           =   4
         End
         Begin VB.Menu mnuESendToWhat 
            Caption         =   "Send &Web Address to Clipboard"
            Index           =   5
         End
      End
      Begin VB.Menu mnuESB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add Item"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "&Edit Item"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete Item"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu SBO 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "View Item"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditSearch 
         Caption         =   "Find Item"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort Binder"
         Begin VB.Menu mnuSortMethod 
            Caption         =   "Sort by "
            Index           =   0
         End
         Begin VB.Menu mnuSortMethod 
            Caption         =   "Sort by "
            Index           =   1
         End
      End
      Begin VB.Menu EB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditBookmarks 
         Caption         =   "View Bookmarks"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuERecycle 
         Caption         =   "&Open Waste Bin"
      End
   End
   Begin VB.Menu mnuQuick 
      Caption         =   "&Quick List"
      Begin VB.Menu mnuQAdd 
         Caption         =   "&Add Binder to Quick List"
      End
      Begin VB.Menu mnuQEdit 
         Caption         =   "&Edit Quick List"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuQSB 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "0.  mnuQuickLink"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "1.  mnuQuickLink"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "2.  mnuQuickLink"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "3.  mnuQuickLink"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "4.  mnuQuickLink"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "5.  mnuQuickLink"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "6.  mnuQuickLink"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "7.  mnuQuickLink"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "8.  mnuQuickLink"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "9. mnuQuickLink"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "10. mnuQuickLink"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "11. mnuQuickLink"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "12. mnuQuickLink"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "13mnuQuickLink"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "14mnuQuickLink"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "15mnuQuickLink"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "16mnuQuickLink"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "17mnuQuickLink"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "18mnuQuickLink"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "19mnuQuickLink"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickLink 
         Caption         =   "20mnuQuickLink"
         Index           =   20
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      HelpContextID   =   6
      Begin VB.Menu mnuODragDrop 
         Caption         =   "Drag and Drop"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOSB1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOChangeDescription 
         Caption         =   "&Change Description..."
      End
      Begin VB.Menu mnuOChangeActiveTemplate 
         Caption         =   "Change &Active Template"
      End
      Begin VB.Menu mnuOFields 
         Caption         =   "User Defined &Fields..."
      End
      Begin VB.Menu mnuOSB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOPref 
         Caption         =   "&Prefrences"
      End
      Begin VB.Menu mnuOSB3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOUpdate 
         Caption         =   "&Update Codes"
         Begin VB.Menu MnuOUpdateCodes 
            Caption         =   "Update &Zip Codes"
            Index           =   0
         End
         Begin VB.Menu MnuOUpdateCodes 
            Caption         =   "Update &Area Codes"
            Index           =   1
         End
      End
      Begin VB.Menu cmdOEMailMailingList 
         Caption         =   "&Generate E-Mail Mailing List"
      End
      Begin VB.Menu mnuOSB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuODefaultFolder 
         Caption         =   "Default Binder"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   7
      Begin VB.Menu mnuHContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHWhatsNew 
         Caption         =   "What's New"
      End
      Begin VB.Menu mnuHSB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTipoftheDay 
         Caption         =   "&Tip of the Day..."
      End
      Begin VB.Menu mnuHSuggest 
         Caption         =   "Contacting &Us"
      End
      Begin VB.Menu mnuHSB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HelpTips As Integer
Dim HelpTipOpen As Integer
Dim LastButton As Integer
Dim bttn As Double, bttnPush As Integer

Dim ProgramIndex() As Integer

Const HIGHLIGHT = &H8000000D
Const background = &H80000005
' Modified: 2-11-1999
Private Sub AddI()
    Dim cnt As Integer
    If Trim$(Pd.AName) <> "" Or Trim$(Pd.Company) <> "" Then
        PutPD FreeNum, ItemCount + 2, True
        LoadItemsIntoGrid flexItems
    End If
    
    'lstNames.SetFocus
    'lstNames.ListIndex = lstNames.ListCount - 1
    BytesUsed
End Sub

Private Sub AppendToBuffer()
 Dim answer As Integer, item As Integer, i As Integer, ItemDes As String
 item = CurrentIndex
 GetPD FreeNum, item, True
 ItemDes = Trim$(Trim$(Pd.Company) & "  " & Trim$(Pd.AName))

 ' Check to see if item is already in the buffer
 For i = 0 To UBound(PrintBuffer)
    If PrintBuffer(i) = item Then
       MsgBox "'" & ItemDes & "' is already in the print buffer", MB_ICONEXCLAMATION, "Item in Buffer"
       Exit Sub
    End If
 Next i
 answer = MsgBox("Add '" & ItemDes & "' to the print buffer?", MB_ICONQUESTION + MB_YESNO, "Add to Buffer")
 If answer = IDYES Then
      If PrintBuffer(0) <> 0 Then ReDim Preserve PrintBuffer(UBound(PrintBuffer) + 1)
      PrintBuffer(UBound(PrintBuffer)) = item   ' Add item to buffer
      UpdatePrintBuffer
 End If
End Sub






Private Sub cmdOEMailMailingList_Click()
    frmEMailList.Show
End Sub

Private Sub flexItems_DblClick()
    If flexItems.Row >= 1 Then ViewItem CurrentIndex
End Sub

Private Sub flexItems_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    If button = RIGHT_DOWN And FolderOpen Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub flexItems_SelChange()
    Dim newRow As Integer
    
    If flexItems.Row > 0 Then
        newRow = flexItems.Row + 1
    
        ShowUnselectedRow flexItems, CurrentIndex - 1
    
        CurrentIndex = newRow
    
        ShowSelectedRow flexItems, CurrentIndex - 1
        BytesUsed
        BookmarkState
    
        ' Show the record if the view window is already open
        If ViewWindowOpen Then ShowRecord CurrentIndex
    End If
End Sub

Private Sub Form_Activate()
    Refresh
End Sub

Private Sub Form_Click()
    'TestSize.Show
    If txtTitle.Visible Then txtTitle_LostFocus
End Sub


Private Sub Form_Load()
    Dim DFolder As String, Fetch As String, FileName As String
    Dim x As Integer
    CloseFolder
    'On Error Resume Next
    CenterMe Me
    Screen.MousePointer = 11
    CurrentCaption = SPROGRAMNAME
    
    '----------------------------------------------------------------
    'Set Startup values
    '----------------------------------------------------------------
    ChangeControlState False
    SetCRLF
    SetCript 3     'Set Encription Key to default
    ChangeControlState False
    lblNumItems = "0"
    lblTime = CStr(Time)
    '----------------------------------------------------------------
    'Get the users name from the WIN.INI file
    '----------------------------------------------------------------

    UserInformation.UName = GetMyIni("ERSOFT.INI", "User", "Name")
    UserInformation.UCompany = GetMyIni("ERSOFT.INI", "User", "Company")
    If (UserInformation.UName = "") Then
        Unload SplashScreen
        Screen.MousePointer = vbNormal
        FirstTimeUser.Show 1
    End If
    ' Get Options from file

    Fetch$ = GetMyIni("BDMAN.INI", "BDMAN", "HelpTips")
    HelpTips = Val(Trim$(Fetch$))
    'mnuOpt.Checked = HelpTips

    Fetch = GetMyIni("BDMAN.INI", "BDMAN", "DragDrop")
    DragDrop = Val(Fetch)
    mnuODragDrop.Checked = DragDrop

    'SetAppHelp (Me.hwnd), "bdman.hlp" 'To setup helpfile variables
    
'---------------------------------------------------------------
'Final Varible startup
'---------------------------------------------------------------
    StartNewForm = True
    imgBinderHeader.Picture = Image1.Picture
    
    If NoteBook = 0 Then
        'If a file name is used has a parametor open it
        FileName$ = Command$
        FileName$ = Trim$(FileName$)
        If FileName$ <> "" Then
            OpenFolder FileName$
        Else
            DFolder = GetMyIni("BDMAN.INI", "Recent Files", "DefFolder")
            If Dir$(DFolder) = "" Then
                SetMyIni "BDMAN.INI", "Recent Files", "DefFolder", ""
            Else
                OpenFolder DFolder
                mnuODefaultFolder.Checked = True
            End If
        End If
 
    End If
    ReDim PrintBuffer(0)
    PrintBuffer(0) = 0
    lblStatus = ""
    
    InitializeGrid flexItems
    CurrentIndex = 2
End Sub



Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    HideTip
    EraseHighlightedButton 0, bttn
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        ' Check for minimal height and width
        lblStatus = ""
        If Height < 4470 Then
            Height = 4470
            lblStatus = "Height too small; enlarged to 4470 twips."
        End If
        If Width < 4530 Then
            Width = 4530
            If lblStatus <> "" Then lblStatus = lblStatus + "  "
            lblStatus = lblStatus + "Width too small; enlarged to 4530 twips."
        End If
        lneMenuBar.X2 = Me.ScaleWidth
       
        
        ' Adjust the Width of the Tool List and Item List
        picToolbar.Width = ScaleWidth - picToolbar.Left
        picListTools.Width = ScaleWidth - picListTools.Left
        flexItems.Width = ScaleWidth - flexItems.Left
        
        ' Adjust the Location of the tools list and the height of the item list
        flexItems.Height = ScaleHeight - flexItems.Top - picStatusBar.Height - picListTools.Height - 100
        picListTools.Top = flexItems.Top + flexItems.Height + 50
        
        ' Move Items to the right of the toolbar
        imgRecycle.Left = picListTools.ScaleWidth - 500
        imgNotebook.Left = imgRecycle.Left - 20 - imgRecycle.Width
        imgPrintBuffer.Left = imgNotebook.Left - 120 - imgNotebook.Width
        
        InitializeGrid flexItems
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Close #TemplateNumber
 Close #FreeNum
 If BigIcon Then
    SetMyIni "BDMAN.INI", "Big Icon", "Top", Str$(frmBigIcon.Top)
    SetMyIni "BDMAN.INI", "Big Icon", "Left", Str$(frmBigIcon.Left)
 End If
    
 If NoteBook = 1 Then
    x% = MsgBox("Close the Notebook also?", MB_YESNO + MB_ICONINFORMATION, SPROGRAMNAME)
    If x% = IDYES Then
      Unload frmMDI
      If NoteBook = 1 Then  'User canceled unload of notebook
         Cancel = True
         Exit Sub
      End If
      End
    Else
      NoteBook = 2
      Exit Sub
    End If
 End If
 End
End Sub

Private Sub HideTip()
If HelpTips = False Then Exit Sub
If HelpTipOpen Then
  'Me.Caption = CurrentCaption
  BytesUsed
  HelpTipOpen = False
End If
End Sub


Private Sub imgAddButton_Click()
    AddItem     ' Start Add Item process
End Sub

Private Sub imgBinderHeader_Click()
    If txtTitle.Visible Then txtTitle_LostFocus
End Sub

Private Sub imgBinderHeader_DblClick()
    lblFolder_DblClick
End Sub


Private Sub imgBinderHeader_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If FolderOpen Then ShowTip imgBinderHeader
End Sub


Private Sub imgBookmarkButton_Click()
  EraseHighlightedButton 0, bttn
  HideTip
  LoadBookmarks
End Sub

Private Sub imgBookmarkButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgBookmarkButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgBookmarkButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.4 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 1.4
    End If
    Controlin3DonPicbox imgBookmarkButton, 1, 0, bttnPush, picToolbar
    ShowTip imgBookmarkButton
End Sub

Private Sub imgBookmarkButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgBookmarkButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgDeleteButton_Click()
    DeleteItem      ' Start delete process
End Sub

Private Sub imgEditButton_Click()
    EditItem    ' Start edit process
End Sub

Private Sub imgEditSortButton_Click()
 EraseHighlightedButton 0, bttn
 HideTip
 mnuSortMethod(0).Caption = "Sort by " & UserField(0)
 mnuSortMethod(1).Caption = "Sort by " & UserField(1)
 PopupMenu mnuSort
End Sub

Private Sub imgEditSortButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgEditSortButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgEditSortButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.3 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 1.3
    End If
    
    Controlin3DonPicbox imgEditSortButton, 1, 0, bttnPush, picToolbar
    ShowTip imgEditSortButton
End Sub

Private Sub imgEditSortButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgEditSortButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileNewButton_Click()
    EraseHighlightedButton 0, bttn
    HideTip
    Call mnuFileNew_CLick
End Sub

Private Sub imgFileNewButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgFileNewButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileNewButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.1 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
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
 EraseHighlightedButton 0, bttn
 HideTip
 Call mnuFileOpen_CLick
End Sub

Private Sub imgFileOpenButton_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgFileOpenButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFileOpenButton_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.2 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 1.2
    End If
    Controlin3DonPicbox imgFileOpenButton, 1, 0, bttnPush, picToolbar
    ShowTip imgFileOpenButton
End Sub

Private Sub imgFileOpenButton_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgFileOpenButton_MouseMove button, Shift, x, y
End Sub

Private Sub imgFindButton_Click()
    frmLocateIt.Show
End Sub

Private Sub imgListDown_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgListDown_MouseMove button, Shift, x, y
End Sub

Private Sub imgListDown_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 4.2 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 4.2
    End If
    Controlin3DonPicbox imgListDown, 1, 0, bttnPush, picListTools
    ShowTip imgListDown
End Sub

Private Sub imgListDown_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgListDown_MouseMove button, Shift, x, y
End Sub

Private Sub imgListUp_Click()
    Dim i As Integer
    
    EraseHighlightedButton 0, bttn
    HideTip
    
    i = CurrentIndex
    ShowUnselectedRow flexItems, CurrentIndex
    If i = 2 Then
        MsgBox "The selected item is already the topmost item in this binder.", MB_ICONINFORMATION, prog
    Else
        Swap i, i - 1, FreeNum
        LoadList
        CurrentIndex = i - 1
    End If
    ShowSelectedRow flexItems, CurrentIndex - 1
End Sub

Private Sub imgListUp_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgListUp_MouseMove button, Shift, x, y
End Sub


Private Sub imgListUp_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 4.1 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 4.1
    End If
    Controlin3DonPicbox imgListUp, 1, 0, bttnPush, picListTools
    ShowTip imgListUp
End Sub


Private Sub imgListUp_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgListUp_MouseMove button, Shift, x, y
End Sub

Private Sub imgNotebook_DblClick()
 ' Open the notebook
 Screen.MousePointer = 11
 If frmMDI.WindowState = 1 Then frmMDI.WindowState = 0
 frmMDI.Show   'Load the Notebook
 Screen.MousePointer = 0
End Sub

Private Sub imgNotebook_DragDrop(source As Control, x As Single, y As Single)
 SendToNotebook
End Sub

Private Sub imgNotebook_DragOver(source As Control, x As Single, y As Single, State As Integer)
 'lstNames.DragIcon = splashscreen.imgDrop.Picture
End Sub

Private Sub imgNotebook_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
ShowTip imgNotebook
End Sub

Private Sub imgPlaceBookmark_Click()
    EraseHighlightedButton 0, bttn
    HideTip
    
    GetPD FreeNum, CurrentIndex, False
    If Left$(imgPlaceBookmark.Tag, 1) = "R" Then
        ' Remove Bookmark
        Pd.Bookmark = False
    Else
        ' Add Bookmark
        Pd.Bookmark = True
    End If
    PutPD FreeNum, CurrentIndex, False
    BookmarkState
End Sub

Private Sub imgPlaceBookmark_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgPlaceBookmark_MouseMove button, Shift, x, y
End Sub

Private Sub imgPlaceBookmark_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 4.3 Then
        EraseHighlightedButton 0, bttn
        bttnPush = False
        bttn = 4.3
    End If
    Controlin3DonPicbox imgPlaceBookmark, 1, 0, bttnPush, picListTools
    ShowTip imgPlaceBookmark
End Sub


Private Sub imgPlaceBookmark_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgPlaceBookmark_MouseMove button, Shift, x, y
End Sub

Private Sub imgPrintBuffer_DblClick()
 Load frmPrintBuffer
 UpdatePrintBuffer
 frmPrintBuffer.Show 1
End Sub

Private Sub imgPrintBuffer_DragDrop(source As Control, x As Single, y As Single)
   AppendToBuffer
End Sub

Private Sub imgPrintBuffer_DragOver(source As Control, x As Single, y As Single, State As Integer)
 'lstNames.DragIcon = splashscreen.imgDrop.Picture
End Sub

Private Sub imgPrintBuffer_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
ShowTip imgPrintBuffer
End Sub

Private Sub imgPrintButton_Click()
    PrintCommand
End Sub

Private Sub imgRecycle_DblClick()
   frmRecycle.Show
End Sub

Private Sub imgRecycle_DragDrop(source As Control, x As Single, y As Single)
 DeleteItem
End Sub

Private Sub imgRecycle_DragOver(source As Control, x As Single, y As Single, State As Integer)
'lstNames.DragIcon = splashscreen.imgDrop.Picture
End Sub

Private Sub imgRecycle_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 ShowTip imgRecycle
End Sub

Private Sub imglistDown_Click()
    On Error Resume Next
    
     EraseHighlightedButton 0, bttn
     HideTip
    
    Dim i As Integer
    i = CurrentIndex
    ShowUnselectedRow flexItems, CurrentIndex - 1
    If i - 2 = ItemCount - 1 Then
        MsgBox "The selected item is already the bottommost item in this binder.", MB_ICONINFORMATION, prog
    Else
        Swap i, i + 1, FreeNum
        LoadList
        CurrentIndex = CurrentIndex + 1
    End If
    ShowSelectedRow flexItems, CurrentIndex - 1
End Sub

Private Sub imgViewButton_Click()
    ViewItem CurrentIndex
End Sub

Private Sub lblAddButton_Click()
    AddItem     ' Start Add Item process
End Sub

Private Sub lblDeleteButton_Click()
    DeleteItem      ' Start delete process
End Sub

Private Sub lblEditButton_Click()
    EditItem    ' Start edit process
End Sub

Private Sub lblFindButton_Click()
    frmLocateIt.Show
End Sub

Private Sub lblFolder_Change()
    OpenFolderName = lblFolder.Caption
    OpenFolderName = Trim$(OpenFolderName)
End Sub

Private Sub lblFolder_Click()
    If txtTitle.Visible Then txtTitle_LostFocus
End Sub


Private Sub lblFolder_DblClick()
    If FolderOpen Then
        txtTitle = lblFolder
        txtTitle.SelStart = 0
        txtTitle.SelLength = Len(txtTitle)
        txtTitle.Visible = True
        txtTitle.SetFocus
        lblStatus = "Enter a new description and press ENTER.  Press ESC to cancel."
    End If
End Sub

Private Sub lblFolder_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If FolderOpen Then ShowTip lblFolder
End Sub



Private Sub lblPrintButton_Click()
    PrintCommand
End Sub


Private Sub lblStatus_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nYear As Integer, FrankAge As Integer
    ' This code implements some interesting EasterEggs
    Select Case Format(Date, "mm-dd")
        Case "12-25": lblStatus = "Merry Christmas"
        Case "02-16": lblStatus = "On this date in 1999 this easter egg was added to PDM"
        Case "08-03"
            ' Calculate how old Frank Rietta is
            nYear = Val(Format(Now, "yyyy"))
            FrankAge = nYear - 1982
            lblStatus = "The author of this program just turned " & Format(FrankAge) & " today!"
            
    End Select
End Sub

Private Sub lblViewButton_Click()
    ViewItem CurrentIndex
End Sub

Private Sub mnuECopy_Click()
  Dim msg As String
  Screen.MousePointer = 11
  GetPD FreeNum, CurrentIndex, True
  msg = "`~BDM10@#%~`&"
  msg = msg & Trim$(Pd.Company) & "&"
  msg = msg & Trim$(Pd.AName) & "&"
  msg = msg & Trim$(Pd.Address) & "&"
  msg = msg & Trim$(Pd.City) & "&"
  msg = msg & Trim$(Pd.State) & "&"
  msg = msg & Trim$(Pd.Zip_Code) & "&"
  msg = msg & Trim$(Pd.Home_Phone) & "&"
  msg = msg & Trim$(Pd.Bus_Phone) & "&"
  msg = msg & Trim$(Pd.Pager) & "&"
  msg = msg & Trim$(Pd.Fax) & "&"
  msg = msg & Trim$(Pd.E_Mail) & "&"
  msg = msg & Trim$(Pd.WebPage) & "&"
  msg = msg & Trim$(Pd.Notes) & "&"
  msg = Trim$(msg)
  Clipboard.SetText msg
  Screen.MousePointer = 0
End Sub

Private Sub mnuEdit_Click()
    Dim text As String
    mnuSortMethod(0).Caption = "Sort by " & UserField(0)
    mnuSortMethod(1).Caption = "Sort by " & UserField(1)
    text = Clipboard.GetText()
    text = Trim$(text)
    If Not FolderOpen Then
        mnuECopy.Enabled = False
        mnuEPaste.Enabled = False
    Else
        mnuECopy.Enabled = True
        If Left(text, 13) <> "`~BDM10@#%~`&" Then
            mnuEPaste.Enabled = False
        Else
            mnuEPaste.Enabled = True
        End If
    End If
End Sub

Private Sub mnuEditAdd_Click()
    AddItem     ' Start Add Item process
End Sub

Private Sub mnuEditBookmarks_Click()
LoadBookmarks
End Sub

Private Sub mnuEditDelete_Click()
 DeleteItem  'Call delete procedure in modual1.bas
End Sub

Private Sub mnuEditModify_Click()
    EditItem
End Sub

Private Sub mnuEditSearch_Click()
    frmLocateIt.Show
End Sub

Private Sub mnuEditView_Click()
    ViewItem CurrentIndex
End Sub

Private Sub mnuEPaste_Click()
    Dim text As String, EndPos As Integer, CurFld As Integer
    
    If ItemCount = 100 Then
        MsgBox "Can't add any more items", 48, "Binder Full"
        Exit Sub
    End If

    Screen.MousePointer = 11

    ReDim Fld(12) As String
    text = Clipboard.GetText()
    text = Trim$(text)
    text = right$(text, Len(text) - 13)
    
    'Parse the string
    CurFld = 0
    Do While Len(text) > 0
        nextpos = InStr(text, "&")
        Fld(CurFld) = Left$(text, nextpos - 1)
        text = right$(text, Len(text) - nextpos)
        CurFld = CurFld + 1
    Loop
    
    Pd.Company = Fld(0)
    Pd.AName = Fld(1)
    Pd.Address = Fld(2)
    Pd.City = Fld(3)
    Pd.State = Fld(4)
    Pd.Zip_Code = Fld(5)
    Pd.Home_Phone = Fld(6)
    Pd.Bus_Phone = Fld(7)
    Pd.Pager = Fld(8)
    Pd.Fax = Fld(9)
    Pd.E_Mail = Fld(10)
    Pd.WebPage = Fld(11)
    Pd.Notes = Fld(12)
    
    AddI          'Add Item
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuERecycle_Click()
 frmRecycle.Show
End Sub

Private Sub mnuESendTo_Click()
  Dim v As Integer

  GetPD FreeNum, CurrentIndex, False
  If Trim$(Pd.E_Mail) = "" Then
       mnuESendToWhat(4).Enabled = False
       v = False
  Else
       mnuESendToWhat(4).Enabled = True
       v = True
  End If
  If Trim$(Pd.WebPage) = "" Then
       mnuESendToWhat(5).Enabled = False
       v = False
  Else
       mnuESendToWhat(5).Enabled = True
       v = True
  End If
  'mnuESendToWhat(3).Visible = v   ' Only display seperator if one of the above items are visible
End Sub

Private Sub mnuESendToWhat_Click(Index As Integer)
  Screen.MousePointer = 11
  Select Case Index
      Case 0: AppendToBuffer
      Case 1: Clipboard.SetText NoBlankFields()
      Case 2: SendToNotebook
      Case 3  ' Seperator Bar
      Case 4
         GetPD FreeNum, CurrentIndex, False
         Pd.E_Mail = Decript$(Trim$(Pd.E_Mail))
         Clipboard.SetText Trim$(Pd.E_Mail)
      Case 5
         GetPD FreeNum, CurrentIndex, False
         Pd.WebPage = Decript$(Trim$(Pd.WebPage))
         Clipboard.SetText Trim$(Pd.WebPage)
  End Select
  Screen.MousePointer = 0
End Sub

Private Sub mnuExportSAT_Click()
 Dim TextFile As String, TextFN As Integer, i As Integer
 
 txt = True
 TextFile = GetFile("Save Binder as text", 2)
 txt = False
 If TextFile = "" Then Exit Sub
 Screen.MousePointer = 11
 TextFN = FreeFile
 Load frmSort
 frmSort.Label1 = "Saving Binder as " & TextFile
 frmSort.Image1.Picture = frmSort.Image2.Picture
 frmSort.Show
 DoEvents
 
 Open TextFile For Output As #TextFN
   FolderN$ = lblFolder.Caption
   cnt% = ItemCount  'lstNames.ListCount
   Print #TextFN, "---------------------------------------------------------------------------------------------------------------------"
   Print #TextFN, " Exported from the Personal Directory Manager on " & Date$
   Print #TextFN, "---------------------------------------------------------------------------------------------------------------------"
   
   Print #TextFN, "Name of Binder: " & Trim$(FolderN$)
   Print #TextFN, Trim$(Str$(cnt%)); " items in folder"
   Print #TextFN, ""
   For i = 2 To 101
     GetPD FreeNum, i, True
     
     If Trim$(Pd.Company) <> "" Or Trim$(Pd.AName) <> "" Then
         If Trim(Pd.Company) <> "" Then Print #TextFN, UserField(0); ": "; Pd.Company
         If Trim(Pd.AName) <> "" Then Print #TextFN, UserField(1); ": "; Pd.AName
         If Trim(Pd.Address) <> "" Then Print #TextFN, UserField(2); ": "; Trim$(Pd.Address)
         If (Trim(Pd.City) <> "") Or (Trim(Pd.State) <> "") Or (Trim(Pd.Zip_Code) <> "") Then
            Print #TextFN, UserField(3); ": "; RTrim$(Pd.City); "     "; UserField(4); ": "; RTrim$(Pd.State); "     "; UserField(5); ": "; RTrim$(Pd.Zip_Code)
         End If
         If Trim(Pd.Home_Phone) <> "" Then Print #TextFN, UserField(6); ": "; Trim$(Pd.Home_Phone)
         If Trim(Pd.Bus_Phone) <> "" Then Print #TextFN, UserField(7); ": "; Trim$(Pd.Bus_Phone)
         If Trim(Pd.Pager) <> "" Then Print #TextFN, UserField(8); ": "; Trim$(Pd.Pager)
         If Trim(Pd.Fax) <> "" Then Print #TextFN, UserField(9); ": "; Trim$(Pd.Fax)
         If Trim(Pd.E_Mail) <> "" Then Print #TextFN, UserField(10); ": "; Trim$(Pd.E_Mail)
         If Trim(Pd.WebPage) <> "" Then Print #TextFN, UserField(11); ": "; Trim$(Pd.WebPage)
         If Trim(Pd.Notes) <> "" Then Print #TextFN, UserField(12); ": "; Trim$(Pd.Notes)
         Print #TextFN, crlf
      End If
    Next i
    
 Close #TextFN
 Unload frmSort
 Screen.MousePointer = 0
 x% = MsgBox("Binder Saved as " & TextFile & crlf & crlf & "Do you wish to view the file now?", MB_YESNO + MB_ICONQUESTION, SPROGRAMNAME)
 If x% = IDYES Then
     StartNewForm = False
     frmMDI.Show
     OpenFile (TextFile)
     StartNewForm = True
 End If
End Sub
Private Sub mnuFFolderPathInfo_Click()
    BinderInformationDialog.Show 1
End Sub

Private Sub mnuFileBackup_Click()
 On Error Resume Next 'Backup may be unloaded in form load
 Si = GetFName(sBinderFileName)
 frmBackup.Show 1
End Sub

Private Sub mnuFileClose_Click()
    CloseFolder
End Sub

Private Sub mnuFileExit_Click()
 Unload frmMain
End Sub

Private Sub mnuFileNew_CLick()
    ' Create a New Binder
    ' Prompts for a Binder Description and Template
    ' Dispays a Save As dialog to get the file name for the Binder
    ' Creates, Initializes, and Opens the new binder.
    
    Dim i As Integer, temp  As String, FileName As String
    
    temp$ = OpenFolderName
    OpenFolderName = "New Binder"
    DN$ = NameFolder("New Binder", True)
    OpenFolderName = temp$
    If DN$ = "" Then Exit Sub
 
    ' Get the file name from the user
    FileName = GetFile("Save Binder As", 2)
    If FileName <> "" Then
        Screen.MousePointer = 11
        
        If FolderOpen Then CloseFolder
        
        sBinderFileName = Trim$(LCase$(FileName))
        FreeNum = FreeFile
        Open sBinderFileName For Random As #FreeNum Len = PD_Size
        
        FolderName DN$           'Name Binder
        lblFolder = DN$
     
        GetPD FreeNum, 1, False
        Pd.Home_Phone = "110"        'Defult Prefrences
        PutPD FreeNum, 1, False
     
        ResetPD
        For i = 2 To 101
            PutPD FreeNum, i, True
        Next i
     
        SetTemplateFieldsToActive CurrentFolderTemplate
        SaveFields
     
        ChangeControlState True
        'lstNames.SetFocus
        On Error Resume Next
        'lstNames.ListIndex = 0
        
        imgBinderHeader.Picture = Image2.Picture
        
        FolderOpen = True
        NewCaption sBinderFileName
        
        InitializeGrid flexItems
                        
        Screen.MousePointer = 0
 End If
End Sub

Private Sub mnuFileNotebook_Click(Index As Integer)
If Index = 0 Then
   Screen.MousePointer = 11
   If frmMDI.WindowState = 1 Then frmMDI.WindowState = 2
   frmMDI.Show
   GetRecentFiles
   Screen.MousePointer = 0
ElseIf Index = 1 Then
   Screen.MousePointer = 11
   NewMemo
   Screen.MousePointer = 0
ElseIf Index = 2 Then
   Screen.MousePointer = 11
   Load frmMDI
   Screen.MousePointer = 0
   FOpenProc

End If
End Sub

Private Sub mnuFileOpen_CLick()
    Dim FileName As String
    FileName = GetFile("", 1)
    If FileName <> "" Then OpenFolder FileName$
End Sub

Private Sub mnuFilePrint_Click()
    PrintCommand
End Sub

Private Sub mnuFilePrinterSetup_Click()
x% = BDM_PrinterSetup()
End Sub

Private Sub mnuFMerge_Click()
  'MsgBox "The merge feature is not availiable in this pre-release.", MB_ICONINFORMATION, SPROGRAMNAME
  frmMerge.Show 1
End Sub

Private Sub mnuFRename_Click()
  frmRenameFile.Show 1
End Sub

Private Sub mnuHContents_Click()
    HTMLHelp "index.htm"
End Sub

Private Sub mnuHelpAbout_Click()
    AboutBox.Show 1
End Sub
Private Sub mnuHSuggest_Click()
    HTMLHelp "contact.htm"
End Sub

Private Sub mnuHTipoftheDay_Click()
  TipOfTheDay.Show 1
End Sub



Private Sub mnuHUsersGuide_Click()
    HTMLHelp "usersguide\usersguide.htm"
End Sub

Private Sub mnuHWhatsNew_Click()
    HTMLHelp "whatsnew.htm"
End Sub

Private Sub mnuOChangeActiveTemplate_Click()
    ChangeActiveTemplateDialog.Show 1
    If Not Canceled Then InitializeGrid flexItems
End Sub

Private Sub mnuOChangeDescription_Click()
DN$ = NameFolder("Change Description", False)
If DN$ <> "" Then
    FolderName DN$
    lblFolder = DN$
End If
End Sub

Private Sub mnuODefaultFolder_Click()
mnuODefaultFolder.Checked = Not mnuODefaultFolder.Checked
If mnuODefaultFolder.Checked = True Then
   SetMyIni "BDMAN.INI", "Recent Files", "DefFolder", sBinderFileName
   MsgBox sBinderFileName & " has been set as the default Folder and will be opened when you start the Personal Directory Manager.", MB_ICONINFORMATION, SPROGRAMNAME
Else
   SetMyIni "BDMAN.INI", "Recent Files", "DefFolder", ""
End If

End Sub

Private Sub mnuODragDrop_Click()
mnuODragDrop.Checked = Not mnuODragDrop.Checked 'Toggle State
DragDrop = mnuODragDrop.Checked
SetMyIni "BDMAN.INI", "BDMAN", "DragDrop", Str$(DragDrop)
End Sub

Private Sub mnuOFields_Click()
 Screen.MousePointer = 11
 frmFields.Show 1
End Sub

Private Sub mnuOPref_Click()
  frmPrefrences.Show 1
End Sub



Private Sub MnuOUpdateCodes_Click(Index As Integer)
 If Index = 0 Then
     frmUpdateCodes.Show 1
 ElseIf Index = 1 Then
     Screen.MousePointer = 11
     Load frmUpdateCodes
     frmUpdateCodes.Caption = "Update Area Codes"
     frmUpdateCodes.lblCode(0) = "Area Code:"
     frmUpdateCodes.lblCode(1) = "New Area Code:"
     frmUpdateCodes.txtCode = ""
     frmUpdateCodes.txtChange = ""
     frmUpdateCodes.cmdOK.Enabled = False
     Screen.MousePointer = 0
     frmUpdateCodes.Show 1
 End If
End Sub



Private Sub mnuQAdd_Click()
  ' Add the current folder to the quick list
  Dim x As Integer, free  As Integer
  free = -1
  For x = 0 To 9
      If QuickList(x).Description = "" And free = -1 Then free = x
      If QuickList(x).File = sBinderFileName And LCase$(QuickList(x).Description) = LCase$(lblFolder) Then
          MsgBox "This binder is already listed in the quick list", MB_ICONEXCLAMATION, SPROGRAMNAME
          Exit Sub
      ElseIf QuickList(x).File = sBinderFileName And LCase$(QuickList(x).Description) <> LCase$(lblFolder) Then
          free = MsgBox("This binder is listed in the quick list, but the description has been changed.  Do you wish to update the quick list to show the new description of this binder?", MB_ICONQUESTION + MB_YESNO, SPROGRAMNAME)
          If free = IDYES Then
             QuickList(x).Description = Trim$(lblFolder)
             SaveQuickList
             ShowQuickList
          End If
          Exit Sub
      End If
  Next x
  If x = -1 Then
      MsgBox "There is no more room in the quick list"
      Exit Sub
  End If
  
  x = MsgBox("Are you sure you want to add '" & lblFolder & "' to the quick list?", MB_ICONQUESTION + MB_YESNO, SPROGRAMNAME)
  If x = IDYES Then
      QuickList(free).Description = lblFolder
      QuickList(free).File = sBinderFileName
      SaveQuickList
      ShowQuickList
  End If
End Sub

Private Sub mnuQEdit_Click()
 frmQuickList.Show
End Sub

Private Sub mnuQuick_Click()
  mnuQAdd.Enabled = FolderOpen
End Sub

Private Sub mnuQuickLink_Click(Index As Integer)
    If Dir$(QuickList(Index).File) = "" Then
        MsgBox QuickList(Index).File & " cannot be found.  The removeable media or network connection may not be present at this time.", MB_ICONEXCLAMATION, SPROGRAMNAME
    Else
        SelectedFile = LCase$(QuickList(Index).File)
        If SelectedFile = LCase$(sBinderFileName) Then
            MsgBox "The file " & """" & SelectedFile & """" & " is already open.", MB_ICONINFORMATION, prog
        Else
            CloseFolder
            OpenFolder SelectedFile
        End If
    End If
End Sub

Private Sub mnuSort_Click()
mnuSortMethod(0).Caption = "Sort by " & UserField(0)
mnuSortMethod(1).Caption = "Sort by " & UserField(1)
End Sub

Private Sub mnuSortMethod_Click(Index As Integer)
  Dim temp As String, ni As Integer
  
  Screen.MousePointer = 11
  temp = Caption
  ChangeControlState False

  ni = ItemCount  'lstNames.ListCount
  Sort Index, ni

  ChangeControlState True
  Caption = temp
  Screen.MousePointer = 0
End Sub

Private Sub picListTools_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 4 Then
        EraseHighlightedButton 0, bttn
        bttn = 4
    End If
    HideTip
End Sub


Private Sub picToolbar_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 0 Then
        EraseHighlightedButton 0, bttn
        'picToolbar.Cls
        bttn = 0
    End If
    HideTip
End Sub

Private Sub picToolbar_Paint()
 LineIn3DonPicBox picToolbar, lneTool1

End Sub



Private Sub SendToNotebook()
  If NoteBook = 0 Then frmMDI.Show
  If frmMDI.WindowState = 1 Then frmMDI.WindowState = 0 'Maximize WIndow to normal size
  FileNew
  GetRecentFiles

  frmMDI.ActiveForm.Text1.text = NoBlankFields()
End Sub

Private Sub ShowTip(Con As Control)
 Dim Tip As String
 
 If Not HelpTips Then Exit Sub
 If Not HelpTipOpen Then CurrentCaption = Me.Caption
 
 Tip = Con.Tag
 
 'imgTipInd.Visible = True
 'lblStatus.Left = 1860
 If lblStatus <> Tip Then lblStatus = Tip

 HelpTipOpen = True
End Sub



Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27
            txtTitle = ""
            txtTitle.Visible = False
        Case 13
            txtTitle_LostFocus
            KeyAscii = 0
    End Select
End Sub


Private Sub txtTitle_LostFocus()
    Dim newBinderName As String
    newBinderName = Trim$(txtTitle)
    If newBinderName <> "" Then
        FolderName newBinderName
        lblFolder = newBinderName
    End If
    txtTitle.Visible = False
End Sub



Public Sub BookmarkState()
    GetPD FreeNum, CurrentIndex, False
    If Pd.Bookmark = True Then
        imgPlaceBookmark.Picture = imgRemoveBookmark.Picture
        imgPlaceBookmark.Tag = "Remove the bookmark from the selected item."
    Else
        imgPlaceBookmark.Picture = imgAddBookmark.Picture
        imgPlaceBookmark.Tag = "Add the selected item to the bookmark list."
    End If
End Sub
' Modified 2-12-1999
Public Sub AddItem()
    Dim x As Integer, cnt As Integer
    
    If ViewWindowOpen Then
        x = MsgBox("You cannot create an Item when the View Window is open.  Close View Window?", MB_ICONEXCLAMATION + MB_YESNO, SPROGRAMNAME)
        If x = IDYES Then
            Unload ViewWindow
        Else
            Exit Sub
        End If
    End If
 
    If ItemCount = 100 Then
        MsgBox "This binder is full and no more items can be added to it.  If you need more room, you will have to create a new binder.", 48, "Binder Full"
        Exit Sub
    End If
    
    ResetPD
    Si$ = "Add Item"
    
    Screen.MousePointer = vbHourglass
 
    Load EditWindow
    
    If Prefrences.FillData Then
        EditWindow.txt_Company = Trim$(DataToFill.Company)
        EditWindow.txt_Name = Trim$(DataToFill.AName)
        EditWindow.txt_BusAddress = Trim$(DataToFill.Address)
        EditWindow.txt_BusCity = Trim$(DataToFill.City)
        EditWindow.txt_BusState = Trim$(DataToFill.State)
        EditWindow.txt_BusZipCode = Trim$(DataToFill.Zip_Code)
        EditWindow.txt_HomePhone = Trim$(DataToFill.Home_Phone)
        EditWindow.txt_Busphone = Trim$(DataToFill.Bus_Phone)
        EditWindow.txt_Pager = Trim$(DataToFill.Pager)
        EditWindow.txt_Fax = Trim$(DataToFill.Fax)
        EditWindow.txt_Email = Trim$(DataToFill.E_Mail)
        EditWindow.txt_WebPage.text = Trim$(DataToFill.WebPage)
        EditWindow.txtNotes.text = Trim$(DataToFill.Notes)
    End If
    
    EditWindow.Show 1
    
    Screen.MousePointer = vbHourglass
 
    AddI          'Add Item
    
    Screen.MousePointer = vbNormal
End Sub
' Created on 2-12-1999
' Code adobted from File/Print
Public Sub PrintCommand()
    ReDim SelectedItems(0) As Integer
    If ItemCount > 0 Then
        Si$ = "1"
        SelectedItems(0) = CurrentIndex
    Else
        Si$ = "2"
    End If
    Screen.MousePointer = 11 'Hourglass
    PrintBinderDialog.Show 1
    ResetPD
End Sub
