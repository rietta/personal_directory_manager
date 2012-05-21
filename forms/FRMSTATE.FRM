VERSION 5.00
Begin VB.Form frmStates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "States"
   ClientHeight    =   2865
   ClientLeft      =   2940
   ClientTop       =   1470
   ClientWidth     =   4530
   ClipControls    =   0   'False
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
   ScaleHeight     =   2865
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
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
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblStateBrev 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alabama, AL"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStates 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listed States/Provinces:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State() As String

Private Sub cmdHelp_Click()
'ShowHelpTopic 12
End Sub

Private Sub Command1_Click()
  l% = List1.ListIndex
  SelectedState = State(l%)
  Unload frmStates
End Sub

Private Sub Command2_Click()
 SelectedState = ""
 Unload frmStates
End Sub

Private Sub Form_Load()
 Dim fn As Integer, l As Integer, i As Integer, CurrentLine As String
 Dim File As String
 File = App.Path
 If right$(File, 1) <> "\" Then File = File & "\"
 File = File & "states.dat"
 '------------------------------------------------------------
 'Open states data file and load state names and abreviations
 '------------------------------------------------------------
 fn = FreeFile
 On Error Resume Next
 If Dir$(File) = "" Then
    MsgBox "Can't open STATES.DAT" & crlf & "Rerun Setup to install missing file(s)", MB_ICONEXCLAMATION, SPROGRAMNAME
    Unload frmStates
    Screen.MousePointer = 0
    Exit Sub
 End If
 Open File$ For Input As #fn
 i = 0
 Do While Not EOF(fn) And i <> 60
     Line Input #fn, CurrentLine
     CurrentLine = Trim$(CurrentLine)
     If Left$(CurrentLine, 1) <> "'" And CurrentLine <> "" Then
         l = InStr(CurrentLine, ",")
         List1.AddItem Left$(CurrentLine, l - 1)
         ReDim Preserve State(i) As String
         State(i) = Trim$(right$(CurrentLine, Len(CurrentLine) - l))
         i = i + 1
     End If
 Loop
 Close #fn
 List1.ListIndex = 0
 lblStateBrev = List1.text & ", " & State(0)
 CenterMe Me
 Screen.MousePointer = 0
End Sub

Private Sub List1_Click()
  lblStateBrev = List1.text & ", " & State(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
  l% = List1.ListIndex
  SelectedState = State(l%)
  Unload frmStates
End Sub

