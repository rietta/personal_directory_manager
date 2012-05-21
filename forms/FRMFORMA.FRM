VERSION 5.00
Begin VB.Form frmFormats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time and Date"
   ClientHeight    =   3330
   ClientLeft      =   1800
   ClientTop       =   2520
   ClientWidth     =   3945
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
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
   ScaleHeight     =   3330
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox lstFormats 
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
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Availible Formats:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
SelFormat.SFormat = ""
SelFormat.SIndex = 0
Unload Me
End Sub

Private Sub cmdOK_Click()
SelFormat.SFormat = lstFormats.Text
SelFormat.SIndex = lstFormats.ListIndex
Unload Me
End Sub

Private Sub Form_Load()
 CenterMe Me
 Dim n
 n = Now
 lstFormats.AddItem n
 lstFormats.AddItem Format$(n, "ttttt mmm d, yyyy")
 lstFormats.AddItem Format$(n, "mmmm d, yyyy")
 lstFormats.AddItem Format$(n, "mmm d, yyyy")
 lstFormats.AddItem Format$(n, "dddd, mmmm d, yyyy")
 lstFormats.AddItem Format$(n, "dddd, mmm d, yyyy")
 lstFormats.AddItem Format$(n, "dddd, mmmm d, yyyy ttttt")
 lstFormats.AddItem Format$(n, "dddd, mmm d, yyyy ttttt")
 lstFormats.AddItem Format$(n, "dd-mmmm-yy")
 lstFormats.AddItem Date$
 lstFormats.AddItem Format$(n, "m/d/yy")
 lstFormats.AddItem Format$(n, "ttttt")

lstFormats.ListIndex = 0
End Sub


Private Sub lstFormats_DblClick()
  Call cmdOK_Click
End Sub

