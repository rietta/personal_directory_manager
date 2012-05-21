VERSION 5.00
Begin VB.Form frmLocateIt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locate It"
   ClientHeight    =   5385
   ClientLeft      =   1605
   ClientTop       =   1140
   ClientWidth     =   6570
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
   Icon            =   "SI.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5385
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "View"
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
      Left            =   5520
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   5295
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Notes"
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
      Index           =   12
      Left            =   2880
      TabIndex        =   18
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Case Sensitive"
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
      Left            =   2880
      TabIndex        =   20
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Search All Fields"
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
      Left            =   2880
      TabIndex        =   19
      Top             =   1800
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Web Page"
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
      Index           =   11
      Left            =   2880
      TabIndex        =   17
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "E-Mail"
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
      Index           =   10
      Left            =   2880
      TabIndex        =   16
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Pager"
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
      Index           =   9
      Left            =   2880
      TabIndex        =   15
      Top             =   720
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Fax"
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
      Index           =   8
      Left            =   2880
      TabIndex        =   14
      Top             =   480
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Business Phone"
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
      Index           =   7
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Home Phone"
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
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Zip"
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
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "State"
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
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "City"
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
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Address"
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
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Name"
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
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.OptionButton optFields 
      Caption         =   "Company"
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
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdState 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Select State"
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
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   5295
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Find"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company:                             Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   2760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   2760
      X2              =   6600
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   2760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   2760
      Y1              =   360
      Y2              =   2640
   End
   Begin VB.Label lblSearchIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search in These Fields:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search For:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmLocateIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SM As Integer
Dim searchRes As Integer

Private Sub cmdCancel_Click()
  Si = ""
  Unload Me
End Sub

Private Sub cmdFields_Click()
 cmdState.Visible = False
 Text1.MaxLength = 40
 Text1.SetFocus
 SearchMode = 1
End Sub

Private Sub cmdFind_Click()
  Dim i As Integer
  If Trim$(Text1) = "" Then Exit Sub
  SearchData.SearchMode = SM
  SearchData.CaseSens = chkCase.value
  SearchData.searchText = Trim$(Text1)
  If Height <> 5790 Then
    Height = 5790
    searchRes = True
    CenterMe Me
  End If
  Search
End Sub

Private Sub cmdState_Click()
 Dim IP As String
 If cmdState.Caption = "&Select State" Then
   Screen.MousePointer = 11
   frmStates.Show 1
   If SelectedState <> "" Then Text1 = Trim$(SelectedState)
 Else
    IP = GetIP()
    If IP <> "" Then Text1.SelText = IP
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.SetFocus
 End If

End Sub

Private Sub cmdView_Click()
    cnt% = List1.ListIndex + 2
    CurrentIndex = SRIndex(cnt%)
    ShowRecord CurrentIndex
End Sub

Private Sub Form_Load()
 Screen.MousePointer = 11
 DoTitle lblTitle
 SearchResultsOpen = True
 
 Height = 4215
 CenterMe Me
 SetTabs Me, Me.List1
 Dim i As Integer
 For i = 0 To 12
    If Trim$(UserField(i)) <> "" Then
       optFields(i).Caption = UserField(i)
    Else
       optFields(i).Visible = False
    End If
 Next i

 Text1.text = SearchData.searchText
 If SearchData.SearchMode = 13 Then
     optAll.value = True
 Else
     SM = SearchData.SearchMode
     If Trim$(UserField(SM)) <> "" Then
         optFields(SM).value = True
     Else
         optAll.value = True
     End If
 End If
 If SM = 4 Then
    cmdState.Caption = "&Select State"
    cmdState.Visible = True
 ElseIf SM = 10 Then
    cmdState.Caption = "&Select Service Provider"
    cmdState.Visible = True
 Else
    cmdState.Visible = False
 End If
 chkCase.value = SearchData.CaseSens
 Text1.SelStart = 0
 Text1.SelLength = Len(Text1)
 Screen.MousePointer = 0
 cmdFind.Enabled = False
End Sub

Private Sub Form_Paint()
   LineIn3D Line4
   LineIn3D Line1
   LineIn3D Line2
   LineIn3D Line3
End Sub

Private Sub List1_DblClick()
    cmdView_Click
End Sub

Private Sub optAll_Click()
  SM = 13
  cmdState.Visible = False
End Sub

Private Sub optFields_Click(Index As Integer)
 SM = Index
 If SM = 4 Then
    cmdState.Caption = "&Select State"
    cmdState.Visible = True
 ElseIf SM = 10 Then
    cmdState.Caption = "&Select Service Provider"
    cmdState.Visible = True
 Else
    cmdState.Visible = False
 End If
End Sub

Private Sub Option1_Click()
cmdState.Visible = True
Text1.text = ""
Text1.MaxLength = 2
Text1.SetFocus
SearchMode = 2
End Sub

Private Sub Text1_Change()
    If Trim$(Text1) = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

