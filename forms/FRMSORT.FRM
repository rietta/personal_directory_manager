VERSION 5.00
Begin VB.Form frmSort 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   1245
   ClientLeft      =   3315
   ClientTop       =   3480
   ClientWidth     =   3540
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
   ScaleHeight     =   1245
   ScaleWidth      =   3540
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   120
      Picture         =   "FRMSORT.frx":0000
      Top             =   1320
      Width           =   765
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   120
      Picture         =   "FRMSORT.frx":066A
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "One moment please..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sorting Binder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterMe Me
StayOnTop Me
End Sub

