VERSION 2.00
Begin Form frmSearchResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Results"
   ClientHeight    =   3720
   ClientLeft      =   1800
   ClientTop       =   1875
   ClientWidth     =   6210
   ClipControls    =   0   'False
   Height          =   4125
   Icon            =   FRMSEARC.FRX:0000
   Left            =   1740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6210
   Top             =   1530
   Width           =   6330
   Begin ListBox List1 
      Height          =   3150
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4575
   End
   Begin CommandButton cmdBookmarkView 
      Caption         =   "&View"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin CommandButton cmdBookMarkClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company:                             Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End

Sub cmdBookMarkClose_Click ()
 Dim i As Integer
 For i = 2 To 100
     SRIndex(i) = 0
 Next i
 Unload frmSearchResults
End Sub

Sub cmdBookmarkView_Click ()
 'Let the user view an Item attached to a bookmark
 Cnt% = List1.ListIndex + 2
 CurrentIndex = SRIndex(Cnt%)
 ShowRecord
End Sub

Sub Form_Load ()
 On Error Resume Next
 DoTitle Me.Label1
 Fetch$ = GetMyINI("BDMAN.INI", "Results Window", "Top")
 If Fetch$ <> "" Then Top = Val(Fetch$)
 Fetch$ = GetMyINI("BDMAN.INI", "Results Window", "Left")
 If Fetch$ <> "" Then Left = Val(Fetch$)
 SearchResultsOpen = True
 SetTabs Me, Me.List1
 Show
 DoEvents
End Sub

Sub Form_Paint ()
lblMatches = Str$(List1.ListCount) & " Matches Found"
End Sub

Sub Form_Unload (Cancel As Integer)
On Error Resume Next
SetMyIni "BDMAN.INI", "Results Window", "Top", Str$(Top)
SetMyIni "BDMAN.INI", "Results Window", "Left", Str$(Left)
SearchResultsOpen = False
End Sub

Sub List1_DblClick ()
 Call cmdBookmarkView_Click
End Sub

