VERSION 5.00
Begin VB.Form Fonts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts"
   ClientHeight    =   2940
   ClientLeft      =   1305
   ClientTop       =   2085
   ClientWidth     =   6930
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFontSize 
      Height          =   1935
      Left            =   2640
      Style           =   1  'Simple Combo
      TabIndex        =   24
      Top             =   360
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set as Default Font"
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
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame frame1 
      Caption         =   "Color"
      Height          =   1095
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   22
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   21
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   20
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   16
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorOp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstFonts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Availiable Fonts"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldColor          As Integer

Private Sub Check1_Click()
 If Check1.value = 1 Then
    SetAsDefault = True
 Else
    SetAsDefault = False
 End If
End Sub

Private Sub ColorOp_Click(Index As Integer)
 On Error GoTo ColorError
 Label2.ForeColor = QBColor(Index)
 nPrintColor = Index
ColorError:
 Exit Sub
Resume
End Sub

Private Sub Command1_Click()
 SFont$ = lstFonts.text
 Unload Fonts
End Sub

Private Sub Command2_Click()
   nPrintColor = OldColor
   Unload Fonts
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
Dim i As Integer, PTarget As Integer
On Error Resume Next
For PTarget = 0 To (Printer.FontCount - 1)
    lstFonts.AddItem Printer.Fonts(PTarget)
Next PTarget
 'Center on screen
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
For i = 0 To 15
  ColorOp(i).BackColor = QBColor(i)
Next i
lstFonts.ListIndex = 0
lstFonts.SetFocus
Call lstFonts_CLick
Screen.MousePointer = 0
OldColor = nPrintColor
Label2.ForeColor = QBColor(nPrintColor)
End Sub


Private Sub lstFonts_CLick()
On Error GoTo ErrorFont
Label2.FontName = lstFonts.text
Label3 = ""
Command1.Enabled = True
Check1.Visible = True
Exit Sub
ErrorFont:
   Check1.Visible = False
   Label3 = "Can't Display Font"
   Command1.Enabled = False
   Exit Sub
Resume
End Sub

Private Sub lstFonts_DblClick()
 Call Command1_Click
End Sub

