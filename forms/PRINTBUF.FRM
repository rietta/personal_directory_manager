VERSION 5.00
Begin VB.Form frmPrintBuffer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Buffer"
   ClientHeight    =   4545
   ClientLeft      =   2955
   ClientTop       =   1875
   ClientWidth     =   5250
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
   ScaleHeight     =   4545
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Print"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cl&ear"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&View"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Sort"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Remove"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image imgDown 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4680
      Picture         =   "PRINTBUF.frx":0000
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image imgUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4080
      Picture         =   "PRINTBUF.frx":0282
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 items:"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrintBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bttn As Integer

Private Sub cmdClear_Click()
   ' Clear the buffer
   Dim x As Integer
   x = MsgBox("Are you sure you want to erase the contents of the print buffer?  The original items will not be effected.", MB_ICONQUESTION + MB_YESNO, "Clear Buffer")
   If x = IDYES Then
      ReDim PrintBuffer(0)
      PrintBuffer(0) = 0
      UpdatePrintBuffer
   End If
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdHelp_Click()
 NoHelp
End Sub

Private Sub cmdPrint_Click()
 Si = "3"            'Set flag so the print dialog "knows" to highlight the "Print Buffer" scope option
 PrintBinderDialog.Show 1    'Open Print Dialog
 Screen.MousePointer = 0
End Sub

Private Sub cmdRemove_Click()
 ' Remove the selected item from the buffer
 If List1.ListCount = 1 Then
 'If there is only one item in the buffer, just erase the
 'buffer itself to save time.
     List1.Clear
     ReDim PrintBuffer(0)
     UpdatePrintBuffer
     Exit Sub
 End If
 
 Dim i As Integer
 ' shift the items up
 For i = List1.ListIndex To UBound(PrintBuffer) - 1
     PrintBuffer(i) = PrintBuffer(i + 1)
 Next i
 PrintBuffer(UBound(PrintBuffer)) = 0
 ReDim Preserve PrintBuffer(UBound(PrintBuffer) - 1)
 UpdatePrintBuffer

End Sub

Private Sub cmdSort_Click()
  PopupMenu frmMain.mnuSort
  UpdatePrintBuffer
End Sub

Private Sub cmdView_Click()
 Screen.MousePointer = 11
 Dim item As Integer
 item = List1.ListIndex
 CurrentIndex = PrintBuffer(item)
 ViewWindow.cmdBookmark.Visible = False
 ViewWindow.cmdDelete.Visible = False
 ViewWindow.cmdPrint.Visible = False
 
 ViewWindow.Show 1
 
End Sub

Private Sub Form_Load()
 PrintBufferOpen = True
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 If bttn <> 0 Then
    bttn = 0
    Cls
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 PrintBufferOpen = False
End Sub

Private Sub imgDown_Click()
  ' Shift selected item up
  Dim i As Integer
  i = List1.ListIndex
  If i = UBound(PrintBuffer) Then
     MsgBox "The selected item is already the bottommost item in the print buffer.", MB_ICONINFORMATION, prog
  Else
     VSwap PrintBuffer(i), PrintBuffer(i + 1)
     UpdatePrintBuffer
     List1.ListIndex = i + 1
  End If
End Sub

Private Sub imgDown_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 If bttn <> 2 Then
    bttn = 2
    Cls
 End If
 ControlIn3D imgDown, 2, 0, False
End Sub

Private Sub imgUp_Click()
  ' Shift selected item up
  Dim i As Integer
  i = List1.ListIndex
  If i = 0 Then
     MsgBox "The selected item is already the topmost item in the print buffer.", MB_ICONINFORMATION, prog
  Else
     VSwap PrintBuffer(i), PrintBuffer(i - 1)
     UpdatePrintBuffer
     List1.ListIndex = i - 1
  End If
End Sub

Private Sub imgUp_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 If bttn <> 1 Then
    bttn = 1
    Cls
 End If
 ControlIn3D imgUp, 2, 0, False
End Sub

Private Sub List1_DblClick()
 cmdView_Click
End Sub

Private Sub List1_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 If bttn <> 0 Then
    bttn = 0
    Cls
 End If

End Sub

