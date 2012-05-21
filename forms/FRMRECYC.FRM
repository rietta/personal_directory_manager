VERSION 5.00
Begin VB.Form frmRecycle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waste Bin"
   ClientHeight    =   3705
   ClientLeft      =   3150
   ClientTop       =   2940
   ClientWidth     =   5280
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
   HelpContextID   =   8
   Icon            =   "FRMRECYC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   5280
   Begin VB.Timer tmrFlame 
      Interval        =   1000
      Left            =   4200
      Top             =   0
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delete"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdUndelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Undelete"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "View"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   3240
      Width           =   855
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
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Image imgCanNoFire 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3600
      Picture         =   "FRMRECYC.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCanFire 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3120
      Picture         =   "FRMRECYC.frx":074C
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCan 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4680
      Picture         =   "FRMRECYC.frx":0B8E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblDeletedItems 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deleted Items:"
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
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmRecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
 Unload frmRecycle
End Sub

Private Sub cmdDelete_Click()
' Delete an item from the waste bin.

Dim x As Integer, RecycleFrom As Integer
imgCan.Picture = imgCanFire.Picture
DoEvents
RecycleFrom = List1.ListIndex + 103
Get #FreeNum, RecycleFrom, Pd

Pd.Company = Decript(Trim$(Pd.Company))
Pd.AName = Decript(Trim$(Pd.AName))

x = MsgBox("Deleting an item from the waste bin will, erase all of the data contained in that item." & crlf & crlf & "Do you wish to remove " & Trim$(Pd.Company) & " " & Trim$(Pd.AName) & " from the waste bin?", MB_YESNO + MB_ICONQUESTION, "Delete Item")
If x = IDYES Then
  For i = RecycleFrom To 113    'Shift Items following the deleted item up
     Get FreeNum, i + 1, Pd
     Put FreeNum, i, Pd
  Next i
  LoadRecyclingBin
End If
tmrFlame.Enabled = True
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdUndelete_Click()
'--------------------------------------------------------------------------
'Undelete selected item and recycle it to the user section of the Binder
'--------------------------------------------------------------------------
Dim i As Integer, RecycleFrom As Integer, RecycleTo As Integer

RecycleTo = ItemCount + 2
RecycleFrom = List1.ListIndex + 103

If RecycleTo = 100 Then
   MsgBox "Can't recycle Item because there is no more room in the binder.", MB_ICONEXCLAMATION, "Binder Full"
   Exit Sub
End If


Get #FreeNum, RecycleFrom, Pd
Put #FreeNum, RecycleTo, Pd
'--------------------------------------------------------------------------
'Shift Recycling Bins contents
'--------------------------------------------------------------------------
For i = RecycleFrom To 113    'Shift Items following the deleted item up
   Get FreeNum, i + 1, Pd
   Put FreeNum, i, Pd
Next i
ResetPD
Put #FreeNum, 113, Pd
LoadList
LoadRecyclingBin
End Sub

Private Sub cmdView_Click()
 ViewSelectedItem
End Sub

Private Sub Form_Load()
    SetTabs frmRecycle, frmRecycle.List1
    Show
    DoEvents
    LoadRecyclingBin
End Sub

Private Sub Form_Unload(Cancel As Integer)
RecyclingBinOpen = False
End Sub

Private Sub List1_DblClick()
ViewSelectedItem
End Sub

Private Sub tmrFlame_Timer()
 imgCan.Picture = imgCanNoFire.Picture
 tmrFlame.Enabled = False
End Sub

Private Sub ViewSelectedItem()
 Screen.MousePointer = 11
 temp = Si
 Si = "&H000005&"
 Load ViewWindow
 Si = temp
 ViewWindow.cmdMemo(0).Visible = False
 ViewWindow.cmdMemo(1).Visible = False
 ViewWindow.cmdBookmark.Visible = False
 ViewWindow.cmdModify.Visible = False
 ViewWindow.cmdPrint.Visible = False
 ViewWindow.cmdDelete.Visible = False

 ViewWindow.cmdCopyToClip.Visible = False
 ViewWindow.cmdClose.Top = 4680
 ViewWindow.cmdClose.Left = 7080
 ViewWindow.cmdClose.Width = 1335
 
 'Load Record
 CI% = List1.ListIndex + 103
 GetPD FreeNum, CI%, True
   ViewWindow.Caption = RTrim$(Pd.AName) & "  --  " & RTrim$(Pd.Company)
   ViewWindow.lblDisplay(0) = RTrim$(Pd.Company)
   ViewWindow.lblDisplay(1) = RTrim$(Pd.AName)
   ViewWindow.lblDisplay(2) = RTrim$(Pd.Address)
   ViewWindow.lblDisplay(3) = RTrim$(Pd.City)
   ViewWindow.lblDisplay(4) = RTrim$(Pd.State)
   ViewWindow.lblDisplay(5) = RTrim$(Pd.Zip_Code)
   ViewWindow.lblDisplay(6) = RTrim$(Pd.Home_Phone)
   ViewWindow.lblDisplay(7) = RTrim$(Pd.Bus_Phone)
   ViewWindow.lblDisplay(8) = RTrim$(Pd.Pager)
   ViewWindow.lblDisplay(9) = RTrim$(Pd.Fax)
   ViewWindow.lblDisplay(10) = RTrim$(Pd.E_Mail)
   ViewWindow.lblDisplay(11) = RTrim$(Pd.WebPage)
   ViewWindow.lblDisplay(12) = RTrim$(Pd.Notes)
 
 
 Screen.MousePointer = 0
 ViewWindow.Show 1
End Sub

