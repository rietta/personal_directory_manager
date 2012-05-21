VERSION 5.00
Begin VB.Form AboutBox 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3750
   ClientLeft      =   1590
   ClientTop       =   2010
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   5970
   Begin VB.Frame Frame1 
      Caption         =   "Licensed to:"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4335
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "John Doe"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rietta Software Inc."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Originally Copyrighted ©1996-2000 Rietta Software"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Web-site:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblURL 
      Caption         =   "http://www.rietta.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Directory Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 2.6 for Windows 95, NT, and 98"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblReleaseDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Release 9-15-99"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4605
      TabIndex        =   1
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Index           =   1
      Left            =   120
      Picture         =   "ABOUTBOX.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©2000 Rietta Solutions"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Sub Command1_Click()
 Unload AboutBox
End Sub

Private Sub Form_Load()
    Label5 = UserInformation.UName
    Label6 = UserInformation.UCompany
    lblReleaseDate = RELEASE
    
    lblVersion = "Version:  " & version
    CenterMe Me
End Sub


Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If lblURL.ForeColor <> vbBlue Then lblURL.ForeColor = vbBlue
End Sub

Private Sub lblURL_Click()
    ShellExecute frmMain.hwnd, "open", "http://www.rietta.com", "", "", 0
End Sub

Private Sub lblURL_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If lblURL.ForeColor <> vbRed Then lblURL.ForeColor = vbRed
End Sub
