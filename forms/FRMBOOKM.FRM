VERSION 5.00
Begin VB.Form frmbookmarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bookmarks"
   ClientHeight    =   3390
   ClientLeft      =   1695
   ClientTop       =   2490
   ClientWidth     =   6135
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
   Icon            =   "FRMBOOKM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdBookMarkClose 
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
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
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
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdBookmarkRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Delete"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookmarkView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&View"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company:                             Name:"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmbookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBookMarkClose_Click()
 For i% = 2 To 100
    BMIndex(i%) = False
 Next i%
 Unload frmbookmarks
End Sub

Private Sub cmdBookmarkRemove_Click()
'Removes a bookmark from a Binder Item
inx% = List1.ListIndex + 2
inx% = BMIndex(inx%)
GetPD FreeNum, inx%, True   'Get Record
Pd.Bookmark = False   'Set Bookmark to none
PutPD FreeNum, inx%, True   'Save changes
LoadBookmarks         'Reload bookmark list
End Sub

Private Sub cmdBookmarkView_Click()
 'Let the user view an Item attached to a bookmark
 cnt% = List1.ListIndex + 2
 CurrentIndex = BMIndex(cnt%)
 ShowRecord CurrentIndex
End Sub

Private Sub Form_Load()
 On Error Resume Next
 Fetch$ = GetMyIni("BDMAN.INI", "Bookmark Viewer", "Top")
 If Fetch$ <> "" Then Top = Val(Fetch$)
 Fetch$ = GetMyIni("BDMAN.INI", "Bookmark Viewer", "Left")
 If Fetch$ <> "" Then Left = Val(Fetch$)
 DoTitle Me.Label1
 BookmarksOpen = True
 SetTabs frmbookmarks, frmbookmarks.List1
 Show
 DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SetMyIni "BDMAN.INI", "Bookmark Viewer", "Top", Str$(Top)
SetMyIni "BDMAN.INI", "Bookmark Viewer", "Left", Str$(Left)
BookmarksOpen = False
End Sub

Private Sub List1_DblClick()
Call cmdBookmarkView_Click
End Sub

