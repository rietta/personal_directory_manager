VERSION 5.00
Begin VB.Form frmFileSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Search"
   ClientHeight    =   5250
   ClientLeft      =   585
   ClientTop       =   1440
   ClientWidth     =   7620
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
   Icon            =   "FINDFILE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   7620
   Begin VB.Frame famFiles 
      Caption         =   "File(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   4455
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "*.fob"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search For:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton searchbutton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "directories searched:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim level As Integer
Dim imat(32) As Integer
Dim totalsearched As Integer
Dim savedrive As String
Dim CancelSearch         As Integer

Private Sub cmdHelp_Click()
nohelp
End Sub

Private Sub Command1_Click()
 Unload frmFileSearch
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub DirLoop(thispath As String, thispattern As String)
    If CancelSearch Then Exit Sub
    On Error GoTo FIndFileError
    'get all files on this directory which match pattern
    If right$(thispath, 1) <> "\" Then thispath = thispath + "\"
    Thisfile = Dir$(thispath + thispattern, 0)
    Do While Thisfile <> ""
        List1.AddItem LCase$(thispath + Thisfile)
        Thisfile = Dir$
    Loop

    'get all files on this directory
    'thisfile = Dir$(thispath + "*.", 0)'the "*." pattern will exclude all files without extensions
    Thisfile = Dir$(thispath, 0)
    thesefiles = 0
    ReDim filelist(10)
    Do While Thisfile <> ""
        thesefiles = thesefiles + 1
        If (thesefiles Mod 10) = 0 Then
            ReDim Preserve filelist(thesefiles + 10)
        End If
        filelist(thesefiles) = Thisfile
        Thisfile = Dir$
    Loop

    'now look at all "names" on this directory
    'thisfile = Dir$(thispath + "*.", 16)
    Thisfile = Dir$(thispath, 16) 'the "*." pattern will exclude all files, including directories, with a "." in them
    checkfile = 1
    thesedirs = 0
    ReDim dirlist(10)
    Do While Thisfile <> ""
        'now check to see if this name is on the file list
        If Thisfile = "." Or Thisfile = ".." Then
            'do nothing, we don't count these as directories
        ElseIf Thisfile = filelist(checkfile) Then
            checkfile = checkfile + 1
        Else
            thesedirs = thesedirs + 1
            If (thesedirs Mod 10) = 0 Then ReDim Preserve dirlist(thesedirs + 10)
            dirlist(thesedirs) = Thisfile
        End If
        Thisfile = Dir$
    Loop

    'now lets go through all subdirectories, and call this sub
    For x = 1 To thesedirs
        Call DirLoop(thispath + dirlist(x), thispattern)
    Next x
    totalsearched = totalsearched + x - 1
    Text2.text = totalsearched
    'waitforothers = DoEvents()
Exit Sub
FIndFileError:
  Beep
  MsgBox Error$, MB_ICONEXCLAMATION, SPROGRAMNAME
  CancelSearch = True
  Exit Sub
Resume
End Sub

Private Sub Drive1_Change()
    On Error GoTo handle_nodrive
    Dir1.Path = Drive1.Drive
    savedrive = Drive1.Drive
    Exit Sub
handle_nodrive:
    If Err = 68 Then
        MsgBox Error$, 16, "Error Reading Drive"
        Err = 0
        Drive1.Drive = savedrive
        Resume Next
    Else
        MsgBox "Fatal Error. . . . Quiting"
        End
    End If
End Sub

Private Sub File1_DblClick()
filename$ = Dir1.Path
If right$(filename$, 1) <> "\" Then filename$ = filename$ & "\"
filename$ = filename$ & File1.filename
OpenF filename$
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:\"
    savedrive = Drive1.Drive
    File1.Pattern = Text1.text
    Screen.MousePointer = 0
End Sub

Private Sub List1_DblClick()
filename$ = List1.text
OpenF filename$
End Sub

Private Sub OpenF(filename As String)
Dim i As Integer
For i = 1 To Len(filename)
  If Mid$(filename, i, 1) = "." Then
     EXT$ = right$(filename, Len(filename) - i)
     Exit For
  End If
Next i
EXT$ = UCase$(Trim$(EXT$))
If EXT$ = "TXT" Or EXT$ = "INI" Or EXT$ = "BAT" Or EXT$ = "DOC" Then
   Screen.MousePointer = 11
   If frmMDI.WindowState = 1 Then frmMDI.WindowState = 2
   frmMDI.Show
   OpenFile filename
   Screen.MousePointer = 0
Else
   OpenFolder filename
End If
End Sub

Private Sub searchbutton_Click()
    List1.Clear
    famFiles = "Searching..."
    CancelSearch = False
    Label2.Visible = True
    Text2.Visible = True
    DoEvents
    dpath$ = Dir1.Path
    dpattern$ = Text1.text
    Screen.MousePointer = 11
    totalsearched = 0
    Call DirLoop(dpath$, dpattern$)
    cnt% = List1.ListCount
    famFiles = Str$(cnt%) & " file(s) found:"
    Label2.Visible = False
    Text2.Visible = False
    
    Screen.MousePointer = 0
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    If right(Text1.text, 1) = ";" Then Exit Sub
    File1.Pattern = Text1.text
    If Err And LCase$(Error$) = "invalid property value" Then
        Beep
        MsgBox "Invalid charactor", MB_ICONEXCLAIMATION, "Error"
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim a As String
a = Chr$(KeyAscii)
Select Case a
   Case ":"
      Beep
      KeyAscii = 0
   Case "\"
      Beep
      KeyAscii = 0
   Case "/"
      Beep
      KeyAscii = 0
End Select
End Sub

