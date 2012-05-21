VERSION 5.00
Begin VB.Form frmFiles 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open"
   ClientHeight    =   4035
   ClientLeft      =   285
   ClientTop       =   1905
   ClientWidth     =   9135
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   9135
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
      Height          =   2730
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtFile 
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
      Left            =   720
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboFileTypes 
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
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
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
      Left            =   720
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   8040
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Open"
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
      Left            =   6960
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.ListBox lstLongNames 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Binder Descriptions:"
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
      Height          =   195
      Index           =   4
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drives:"
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
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "List Files:"
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
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   7
      Top             =   3600
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Files:"
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
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Folders:"
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dpattern As String
Dim level As Integer
Dim imat(32) As Integer
Dim totalsearched As Integer

Private Sub cboFileTypes_Change()
   Select Case cboFileTypes.ListIndex
       Case 0: dpattern$ = "*.fo?"
       Case 1: dpattern$ = "*.fob"
       Case 2: dpattern$ = "*.fol"
       Case 3: dpattern$ = "*.*"
   End Select
   Call DirLoop(dpath$, dpattern$)
End Sub

Private Sub cboFileTypes_Click()
 Call cboFileTypes_Change
End Sub

Private Sub cmdCancel_Click()
 SelectedFile = ""
 Unload Me
End Sub

Private Sub cmdOK_Click()
Dim d As String, i As Integer
If cmdOK.Caption = "Save" Then
   d = Trim$(LCase$(txtFile))
   For i = 1 To Len(d)
       Select Case Mid$(d, i, 1)
        Case " ", "*", "?", "/", "|", ",", ";", ":", "+", "=", "[", "]", "(", ")", "&", "^", "<", ">", """"
             MsgBox Trim$(txtFile) & " is an invalid file name.  The file name must be a valid MS-DOS file name and cannot contain any of the following charactors: [space], [comma], *, ?, /, \, |, ., ;, :, +, =, [, ], (, ), &, ^, <, >", MB_ICONEXCLAMATION, "Invalid File Name"
             Exit Sub
       End Select
   Next i
   If right$(dpath$, 1) <> "\" Then dpath$ = dpath$ & "\"
   If Dir$(dpath$ & d) <> "" Then
      i = MsgBox(dpath$ & d & " already exists.   Replace existing file?", MB_ICONEXCLAMATION + MB_YESNO, "Save Binder")
      If i = IDNO Then Exit Sub
   End If
   SelectedFile = dpath$ & d
   LastFileType = cboFileTypes.ListIndex  'Save the index of the selected file type
Else
   SelectedFile = GetFileFromLong()
   LastFileType = cboFileTypes.ListIndex  'Save the index of the selected file type
End If
Unload Me
End Sub

Private Sub Dir1_Change()
    dpath$ = Dir1.Path
    Call DirLoop(dpath$, dpattern$)
End Sub

Private Sub DirLoop(thispath As String, thispattern As String)
    Dim Thisfile As String, thesefiles, checkfile, thesedirs
    Dim x
    Screen.MousePointer = 11
    'On Error GoTo FIndFileError
    lstLongNames.Clear
    'get all files on this directory which match pattern
    If right$(thispath, 1) <> "\" Then thispath = thispath + "\"
    Thisfile = Dir$(thispath + thispattern, 0)
    Do While Thisfile <> ""
        lstLongNames.AddItem LongName(LCase$(thispath), LCase$(Thisfile))
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
    'For x = 1 To thesedirs
    '    Call dirloop(thispath + dirlist(x), thispattern)
    'Next x
    'waitforothers = DoEvents()
    lstLongNames.Enabled = True
    If lstLongNames.ListCount > 0 Then
        lstLongNames.ListIndex = 0
    Else
        lstLongNames.AddItem "<<<no files of the selected type are in this directory>>>"
        lstLongNames.Enabled = False
    End If
    Screen.MousePointer = 0
Exit Sub
FIndFileError:
  Beep
  MsgBox Error$, 48, SPROGRAMNAME
  Screen.MousePointer = 0
  Exit Sub
Resume
End Sub

Private Sub Drive1_Change()
  Dim x As Integer, msg As String
  On Error GoTo DriveError
  Dir1.Path = Drive1.Drive
  savedrive = Drive1.Drive
  ChDrive savedrive         'Change the active drive
  Exit Sub

DriveError:
  If Err = 68 Then
      x = MsgBox("Cannot read from drive " & Drive1.Drive & crlf & "Make sure that the drive door is closed, and that the disk is formatted and free from errors", MB_ICONEXCLAMATION + MB_RETRYCANCEL, "Error Reading Drive " & Drive1.Drive)
      If x = IDRETRY Then
          Resume
      Else
          Drive1.Drive = savedrive
          Dir1.Path = savedrive
          Exit Sub
      End If
  Else
      msg = "The following unexpected error has accured while reading drive " & UCase(Drive1.Drive)
      If right(msg, 1) <> ":" Then
          msg = msg & ": "
      Else
          msg = msg & " "
      End If
      MsgBox msg & crlf & crlf & Error$, MB_ICONEXCLAMATION, "Unexpected Error"
      Drive1.Drive = savedrive
      Exit Sub
  End If
Resume
End Sub

Private Sub Form_Load()
SetCRLF
Dir1.Height = 3180
cboFileTypes.AddItem "All Binders (*.FO?)"
cboFileTypes.AddItem "BDM Binders (*.FOB)"
cboFileTypes.AddItem "PDM Binders (*.FOL)"
cboFileTypes.AddItem "All Files   (*.*)"
cboFileTypes.ListIndex = LastFileType   'Select the item the user last used

If dpath = "" Then
   dpath = Dir1.Path
   savedrive = Left$(dpath, 1)
Else
   Dir1.Path = dpath$
End If
'If SaveDrive = "" Then
'   SaveDrive = Drive1.Drive
'Else
   If savedrive <> "" Then Drive1.Drive = savedrive
'End If
DirLoop dpath$, dpattern$
Screen.MousePointer = 0
End Sub

Private Function GetFileFromLong()
Dim x As Integer, t As String
t = lstLongNames.text
t = Trim$(t)
x = InStr(t, ":") - 1
If x > 0 Then
  If right$(dpath, 1) = "\" Then
      GetFileFromLong = dpath & Trim$(Left$(t, x))
  Else
      GetFileFromLong = dpath & "\" & Trim$(Left$(t, x))
  End If
End If
End Function

Private Function LabelFile(File As String, LongN As String)
Dim temp As String
temp = Trim$(File & ":")
If Len(temp) <= 8 Then
  temp = temp & " " & Chr$(9) & Chr$(9)
ElseIf Len(temp) > 8 Then
  temp = temp & "   " & Chr$(9)
End If
LabelFile = temp & LongN
End Function

Private Function LongName(ByVal Path As String, ByVal File As String)
    Dim BDM As FolderType, fn As Integer, Header  As String
    Dim PDM97 As OldPdType, PDM20 As PDtype
    On Error Resume Next

    File = LCase$(Trim$(File))
    If LCase$(Path & File) = LCase$(sBinderFileName) Then
        LongName = LabelFile(File, "<<<Already Open>>>")
        Exit Function
    End If
    fn = FreeFile
    Open Path & File For Input As #fn
        If Err Then
            LongName = LabelFile(File, "<<<unable to evaluate>>>")
            Exit Function
        End If
        Header = Input$(100, #fn)
        If Err Then
            LongName = LabelFile(File, "<<<unable to evaluate>>>")
            Exit Function
        End If
  
    Close #fn

    Select Case CheckFileFormat(Header)
        Case 0: LongName = LabelFile(File, "<<<not a binder>>>")
        Case 2
            fn = FreeFile
            Open Path & File For Random As #fn Len = BDM_Size
                Get #fn, 1, BDM
                LongName = File & ":" & Chr$(9) & CriptString(Trim$(BDM.WebPage), -3)
            Close #fn
        Case 1
            fn = FreeFile
            Open Path & File For Random As #fn Len = PDM97_Size
                Get #fn, 1, PDM97
                LongName = LabelFile(File, Trim$(PDM97.OFirst_Name))
            Close #fn
        Case -2
            fn = FreeFile
            Open Path & File For Random As #fn Len = PDM20_Size
                Get #fn, 1, PDM20
                LongName = LabelFile(File, Trim$(PDM20.CE_Mail))
            Close #fn
    End Select



End Function

Private Sub lstLongNames_Click()
 Dim t As String, c As String, x As Integer
 If cmdOK.Caption = "Save" Then
    t = lstLongNames.text
    t = right$(t, Len(t) - InStr(t, ":"))
    x = InStr(t, Chr$(9))
    c = ""
    For x = 1 To Len(t)
       If Mid$(t, x, 1) <> Chr$(9) Then c = c & Mid$(t, x, 1)
    Next x
    txtFile = c
    txtFile.SelStart = 0
    txtFile.SelLength = Len(c)
 End If
End Sub

Private Sub lstLongNames_DblClick()
SelectedFile = GetFileFromLong()
LastFileType = cboFileTypes.ListIndex  'Save the index of the selected file type
Unload Me
End Sub

