VERSION 5.00
Begin VB.Form frmBigIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BDM"
   ClientHeight    =   255
   ClientLeft      =   2265
   ClientTop       =   3600
   ClientWidth     =   1560
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
   Icon            =   "BIGICON.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   255
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show BDM"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuFNB 
         Caption         =   "Notebook..."
      End
      Begin VB.Menu mnuFFindFile 
         Caption         =   "Find File..."
      End
      Begin VB.Menu mnuFSB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBigIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Go

Private Sub Command1_Click()
 If Go Then ShowWindows True
End Sub

Private Sub Command1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
 If button = RIGHT_DOWN Then
    Go = False
    PopupMenu mnufile
 Else
    Go = True
 End If
End Sub

Private Sub Form_Load()
 StayOnTop Me
 On Error Resume Next
 Fetch$ = GetMyIni("BDMAN.INI", "Big Icon", "Top")
 If Fetch$ <> "" Then Top = Val(Fetch$)
 Fetch$ = GetMyIni("BDMAN.INI", "Big Icon", "Left")
 If Fetch$ <> "" Then Left = Val(Fetch$)
 BigIcon = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 BigIcon = False
End Sub

Private Sub mnufexit_Click()
        Hide
        Unload frmMain

End Sub

Private Sub mnuFFindFile_Click()
 Screen.MousePointer = 11
 frmFileSearch.Show
End Sub

Private Sub mnuFNB_Click()
        Screen.MousePointer = 11
        If frmMDI.WindowState = 1 Then frmMDI.WindowState = 2
        frmMDI.Show
        Screen.MousePointer = 0
End Sub

