VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SplashScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   1275
   ClientTop       =   2085
   ClientWidth     =   7185
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlog 
      Left            =   1920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   3240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Originally Copyrighted ©1996-2000 Rietta Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Image imgDrag 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "INTRO.frx":0000
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDrop 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   600
      Picture         =   "INTRO.frx":030A
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPrintBufferEnabled 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "INTRO.frx":0614
      Tag             =   "Printer Buffer"
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPrintBufferDisabled 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1200
      Picture         =   "INTRO.frx":091E
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRecycleEn 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3600
      Picture         =   "INTRO.frx":0C28
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRecycleDis 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3120
      Picture         =   "INTRO.frx":106A
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   120
      Picture         =   "INTRO.frx":14AC
      Top             =   720
      Width           =   6225
   End
   Begin VB.Label lblOS 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.62 for Windows 9x and 2000"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblProtected 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "This Program is Protected by United States Copyright Law and International Treaties."
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
      TabIndex        =   2
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Image imgPhone 
      Appearance      =   0  'Flat
      Height          =   2460
      Left            =   6240
      Picture         =   "INTRO.frx":3DEC
      Stretch         =   -1  'True
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©2000 Rietta Solutions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   870
      Left            =   1200
      Picture         =   "INTRO.frx":91D4
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 CenterMe Me

End Sub

Private Sub Timer1_Timer()
    SplashScreenOn = False
    Hide
End Sub
