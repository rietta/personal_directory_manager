VERSION 5.00
Begin VB.Form TipOfTheDay 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3705
   ClientLeft      =   2235
   ClientTop       =   2445
   ClientWidth     =   6600
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
   Icon            =   "TIPOFTHE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   6600
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&About"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdLastTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "<-- &Last Tip"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Next Tip -->"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   5160
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkShowTips 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Tips at Startup"
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
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2865
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Current Tip is Displayed Here>"
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
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1350
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   240
         Picture         =   "TIPOFTHE.frx":030A
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Shape Boarder 
      Height          =   3135
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "TipOfTheDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentTip As Integer

Private Sub chkShowTips_Click()
 Dim f As String
 If chkShowTips.value = 1 Then
    f = "0"
 Else
    f = "1"
 End If
 SetMyIni "BDMAN.INI", "BDMAN", "TipOfTheDay", f
End Sub

Private Sub cmdAbout_Click()
 AboutBox.Show 1
End Sub

Private Sub cmdLastTip_Click()
 If CurrentTip >= 1 Then CurrentTip = CurrentTip - 1
 If CurrentTip = 0 Then cmdLastTip.Enabled = False
 lblTip = GetTip(CurrentTip)
 SetMyIni "BDMAN.INI", "BDMAN", "CurrentTip", Str$(CurrentTip)
End Sub

Private Sub cmdNext_Click()
 Dim temp As String
 CurrentTip = CurrentTip + 1
 temp = GetTip(CurrentTip)
 If temp = "" Then
     CurrentTip = 0
     temp = GetTip(0)
     cmdLastTip.Enabled = False
 Else
     cmdLastTip.Enabled = True
 End If
 lblTip = temp
 SetMyIni "BDMAN.INI", "BDMAN", "CurrentTip", Str$(CurrentTip)
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 Dim Fetch As String
 Fetch = Trim$(GetMyIni("BDMAN.INI", "BDMAN", "TipOfTheDay"))
 If Fetch = "" Or Left$(Fetch, 1) = "0" Then
    chkShowTips.value = 1
 End If
 Fetch = GetMyIni("BDMAN.INI", "BDMAN", "CurrentTip")
 Fetch = Trim$(Fetch)
 CurrentTip = Val(Fetch)
 If CurrentTip > 0 Then Call cmdNext_Click
 If CurrentTip = 0 Then
    cmdLastTip.Enabled = False
    lblTip = GetTip(CurrentTip)
 End If
 CenterMe Me
End Sub

Private Sub Form_Paint()
  ControlIn3D Boarder, 1, 0, True
End Sub

Private Function GetTip(Tip As Integer) As String
Dim msg As String
Select Case Tip
   Case 0: msg = "You can change the titles of the fields in a binder.  Simply select 'User Definible Fields...' from the main window's 'Option' menu."
   Case 1: msg = "You can find the location of the currently open binder by choosing 'Folder Path Information' from the main window's 'File' menu."
   Case 2: msg = "You can make a backup copy of a folder on a diskette by choosing 'Backup Binder' from the main window's 'File' menu."
   Case 3: msg = "You can locate lost files using 'File Search'.  Simply choose 'Find File...' from the main window's 'File' menu."
   Case 4: msg = "You can save a binder as a text file so they can be manipulated in any word processor.  Simply choose 'Save Binder as Text...' from the main window's 'File' menu."
   Case 5: msg = "You can make it easier to quickly open common binders by using the Quick List.  The Quick List works the same way as the Favorites menu does in Internet Explorer.  To add a binder to the Quick List, simply select Add from the Quick List menu."
   Case 6: msg = ""
   Case 7: msg = ""
   Case 8: msg = ""
   Case 9: msg = ""
   Case 10: msg = ""
   Case Else: msg = ""   ' No more tips
End Select
GetTip = msg
End Function

