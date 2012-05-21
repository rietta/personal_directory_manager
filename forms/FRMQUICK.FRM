VERSION 5.00
Begin VB.Form frmQuickList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick List Editor"
   ClientHeight    =   4440
   ClientLeft      =   2115
   ClientTop       =   2160
   ClientWidth     =   6975
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
   Icon            =   "FRMQUICK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   6615
      Begin VB.Label lblFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00000005&
         BackStyle       =   0  'Transparent
         Caption         =   "c:\bdman20\folders\buscon.fob"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Important Business Contacts"
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
         TabIndex        =   8
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Clear"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Open"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
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
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdProperties 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Properties"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1215
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
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Links:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuProp 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuPProp 
         Caption         =   "Item Properties"
      End
      Begin VB.Menu mnuPSB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuPClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmQuickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
 Dim x As Integer, i As Integer
 x = MsgBox("Are you sure you want the delete all of the quick links?", MB_ICONQUESTION + MB_YESNO, SPROGRAMNAME)
 If x = IDYES Then
    For i = 0 To 9
        QuickList(i).Description = ""
        QuickList(i).File = ""
    Next i
    
    SaveQuickList
    ShowQuickList
 End If
End Sub

Private Sub cmdClose_Click()
  
  Unload Me
End Sub

Private Sub cmdOpen_click()
 If FolderOpen Then Close #FreeNum
 OpenFolder QuickList(List1.ListIndex).File
End Sub

Private Sub cmdProperties_Click()
    Dim i As Integer
    i = List1.ListIndex
    QuickLinkEdit.Description = QuickList(i).Description
    QuickLinkEdit.File = QuickList(i).File
    frmQuickLinkProp.Show 1
    QuickList(i).Description = QuickLinkEdit.Description
    QuickList(i).File = QuickLinkEdit.File
    SaveQuickList
    ShowQuickList
    List1.ListIndex = i
End Sub

Private Sub cmdRemove_Click()
 ' Delete the selected link
 Dim i As Integer, C As Integer
 
 ' Shift items up one space and overwrite the selected one
 C = List1.ListIndex
 For i = C To 8
    QuickList(i).Description = QuickList(i + 1).Description
    QuickList(i).File = QuickList(i + 1).File
 Next i
 QuickList(9).Description = ""
 QuickList(9).File = ""
 
 ShowQuickList
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
 SetTabs Me, Me.List1
 QuickListEditorOpen = True
 ShowQuickListEdit
 Show
 Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 QuickListEditorOpen = False
 SaveQuickList
End Sub

Private Sub List1_Click()
 lblDesc = QuickList(List1.ListIndex).Description
 lblFile = QuickList(List1.ListIndex).File
End Sub

Private Sub List1_DblClick()
 cmdOpen_click
End Sub

Private Sub List1_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    If button = 2 Then PopupMenu mnuProp
End Sub

Private Sub mnuPClear_Click()
cmdClear_Click
End Sub

Private Sub mnuPProp_Click()
cmdProperties_Click
End Sub

Private Sub mnuPRemove_Click()
cmdRemove_Click
End Sub

