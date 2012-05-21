VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Dform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   5595
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog dlog 
      Left            =   2880
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   0
   End
   Begin VB.Image imgRecycleDis 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3000
      Picture         =   "DFORM.frx":0000
      Top             =   2040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgRecycleEn 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2400
      Picture         =   "DFORM.frx":01BA
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPrintBufferDisabled 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   480
      Picture         =   "DFORM.frx":05FC
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPrintBufferEnabled 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "DFORM.frx":0906
      Tag             =   "Printer Buffer"
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDrop 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   600
      Picture         =   "DFORM.frx":0C10
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDrag 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "DFORM.frx":0F1A
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Dform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Hide
  Refresh
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Dform
End Sub

