VERSION 5.00
Begin VB.Form frmRenameFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename File"
   ClientHeight    =   1650
   ClientLeft      =   1560
   ClientTop       =   2760
   ClientWidth     =   3975
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
   ScaleHeight     =   1650
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Rename"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtNewName 
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
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblCurrentName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "folder.fob"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Name:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmRenameFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdRename_Click()
  ' Rename the file
  Dim newFName As String
  newFName = GetPath(PDM_F)
  If right$(newFName, 1) <> "\" Then newFName = newFName & "\"
  newFName = newFName & Trim$(txtNewName)
  newFName = LCase$(newFName)
  If newFName = sBinderFileName Then
     MsgBox "You Cannot rename a file with the same name.", MB_ICONEXCLAMATION, SPROGRAMNAME
     txtNewName.SelStart = 0
     txtNewName.SelLength = Len(txtNewName)
     Exit Sub
  ElseIf Dir$(newFName) <> "" Then
     MsgBox newFName & " already exists.  You cannot rename a file with the name of annother file.", MB_ICONEXCLAMATION, SPROGRAMNAME
     txtNewName.SelStart = 0
     txtNewName.SelLength = Len(txtNewName)
     Exit Sub
  End If
  Close #FreeNum
  Name sBinderFileName As newFName    ' Rename file
  Open newFName For Random As #FreeNum Len = PD_Size
  frmMain.Caption = UCase$(GetFName(newFName)) & " - " & prog
  App.title = frmMain.Caption
  Unload Me
End Sub

Private Sub Form_Load()
 lblCurrentName = GetFName(sBinderFileName)
 txtNewName = lblCurrentName
 txtNewName.SelStart = 0
 txtNewName.SelLength = Len(txtNewName)
End Sub



Private Sub txtNewName_Change()
  If Trim$(txtNewName) <> "" Then
     cmdRename.Enabled = True
  Else
     cmdRename.Enabled = False
  End If
End Sub

Private Sub txtNewName_KeyPress(KeyAscii As Integer)
  Select Case Chr$(KeyAscii)
     Case ":", "\", "|", "/", "?", "*", " "
        Beep
        KeyAscii = 0
  End Select
End Sub

