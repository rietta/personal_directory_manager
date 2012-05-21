VERSION 5.00
Begin VB.Form frmWhichAddress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Address Conflict"
   ClientHeight    =   2640
   ClientLeft      =   975
   ClientTop       =   2730
   ClientWidth     =   8115
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   8115
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
      Left            =   7200
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdConvert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Convert"
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
      Left            =   7200
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdConvert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Convert"
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
      Left            =   7200
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bus. Address:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblBAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Business Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label lblHAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Home Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"FRMWHICH.frx":0000
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmWhichAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Dim x As Integer
 x = MsgBox("Are you sure you want to stop converting the binder", MB_ICONQUESTION + MB_YESNO, SPROGRAMNAME)
 If x = IDYES Then
    ConvertAddress = -1
    Unload Me
 End If
End Sub

Private Sub cmdConvert_Click(Index As Integer)
 ConvertAddress = Index
 Unload Me
End Sub

Private Sub Form_Load()
 Dim msg, item
 item = UCase$(Trim$(PD20.CLast_Name) & " " & Trim$(PD20.CFirst_Name))
 
 msg = "The item " & item & " has two address.  The Buisness Directory Manager only supports one address per item.  Which address do you want to convert?"
 Label1 = msg

 msg = Trim$(PD20.cAddress) & crlf
 msg = msg & Trim$(PD20.CCity) & "  " & Trim$(PD20.CState) & ",  " & Trim$(PD20.CZip_Code)
 lblHAddress = msg

 msg = Trim$(PD20.cBus_Address) & crlf
 msg = msg & Trim$(PD20.CBus_City) & "  " & Trim$(PD20.CBus_State) & ",  " & Trim$(PD20.CBus_Zip_Code)
 lblBAddress = msg

End Sub

Private Sub Form_Paint()
    ControlIn3D Label1, 1, 0, False
End Sub

