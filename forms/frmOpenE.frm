VERSION 5.00
Begin VB.Form frmOpenExisting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Existing Directory"
   ClientHeight    =   4995
   ClientLeft      =   735
   ClientTop       =   1200
   ClientWidth     =   8265
   Icon            =   "FRMOPENE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMyFiles 
      Height          =   1620
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   8055
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Open"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdMoreFiles 
      Caption         =   "&More Files"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   1815
      Left            =   4080
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   4
      Top             =   480
      Width           =   4095
      Begin VB.Label lblNumItems 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDirInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Items:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblNumSections 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblDirInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Sections:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "No Information Availiable for This File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblDirInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.ListBox lstRecentFiles 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblHeadLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Binder Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   3
      Top             =   165
      Width           =   4095
   End
   Begin VB.Shape shpHeadLines 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblHeadLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Directory Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2445
      Width           =   7935
   End
   Begin VB.Shape shpHeadLines 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   8055
   End
   Begin VB.Label lblHeadLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   3855
   End
   Begin VB.Shape shpHeadLines 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmOpenExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub


Private Sub cmdMoreFiles_Click()
    Dim moreFile As String
    frmFiles.Show 1
    If SelectedFile <> "" Then Unload Me
End Sub


Private Sub cmdOK_Click()
    Canceled = False
    Unload Me
End Sub


Private Sub Form_Load()
    CenterMe Me
End Sub


Private Sub Label1_Click()

End Sub


