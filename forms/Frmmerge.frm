VERSION 5.00
Begin VB.Form frmMerge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Merge"
   ClientHeight    =   3000
   ClientLeft      =   1470
   ClientTop       =   2640
   ClientWidth     =   5940
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3000
   ScaleWidth      =   5940
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
      Left            =   4680
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Merge Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3375
      Begin VB.OptionButton optAddUpdate 
         Caption         =   "Add and Update Items in Binder"
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
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton optUpdate 
         Caption         =   "Update Items in Binder"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add Items to end of Binder"
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
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   5520
      Picture         =   "Frmmerge.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdSelectSource 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   5520
      Picture         =   "Frmmerge.frx":05C4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblInfoDest 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(source binder will be merged into this file)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblDest 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Binder:"
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
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblSource 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Source Binder:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()

Dim mode As Integer, i As Integer, FileNum As Integer, ff As Integer
Dim sourcefile As String, destfile As String, CheckHeader As String

      sourcefile = lblSource
      destfile = lblDest
      If sourcefile = "" Then
         MsgBox "You must specify a source file.", MB_ICONEXCLAMATION, "Merge"
         cmdSelectSource_click
         Exit Sub
      ElseIf destfile = "" Then
         MsgBox "You must specify a destination file.", MB_ICONEXCLAMATION, "Merge"
         cmdOpen_click
         Exit Sub
      End If
      
      ' Check to see if the source is a valid binder
      FileNum = FreeFile
      
      On Error Resume Next
      
      Open sourcefile For Input As #FileNum
      If Err Then
          MsgBox "Cannot open: " & Trim$(UCase$(sourcefile)), MB_ICONSTOP, SPROGRAMNAME
          Exit Sub
      End If
      CheckHeader = Input$(100, #FileNum)
      If Err = 51 Then GoTo ErrorFormat
      Close #FileNum
      ff = CheckFileFormat(CheckHeader)
      Select Case ff
           Case 0
              GoTo ErrorFormat
           Case -2, 1
              MsgBox GetFileFromPath(sourcefile) & " is a Personal Directory Manager binder." & crlf & crlf & "Because of the big differences between the Personal Directory Manager and Personal Directory Manager binders, they cannot be automatically merged." & crlf & crlf & "If you wish to merge the contents of the two binders, you must first convert the Personal Directory Manager binder to a Personal Directory Manager format.  To convert the binder choose Open from the main windows File menu. You will be prompted to see if you want to convert it, click Yes.", MB_ICONEXCLAMATION, "Incompatible File"
              Exit Sub
      End Select
    Screen.MousePointer = 11
    If optAdd.value Then
        i = MergeBindersAdd(sourcefile, destfile)
        If i = True Then
            If LCase$(PDM_F) = LCase$(sourcefile) Or LCase$(PDM_F) = LCase$(destfile) Then
                OpenFolder PDM_F
            End If
           
        End If
    ElseIf optUpdate.value Then
        'i = MergeBindersUpdate(SourceFile, DestFile)
    ElseIf optAddUpdate.value Then
        'i = MergeBindersAddUpdate(SourceFile, DestFile)
    End If
    If i = 0 Then MsgBox "The Merge of " & GetFileFromPath(sourcefile) & " and " & GetFileFromPath(destfile) & " was unsuccesfull.", MB_ICONEXCLAMATION, "Merge Error"
    Screen.MousePointer = 0
    Unload Me
Exit Sub

ErrorFormat:
 MsgBox sourcefile & " is not a Personal Directory Manager Binder.", MB_ICONEXCLAMATION, "Merge"
 Close #FileNum

End Sub

Private Sub cmdOpen_click()
    Dim FileName As String
    FileName = GetFile("", 1)
    If FileName <> "" Then
        lblDest = LCase$(FileName)
    End If
End Sub

Private Sub cmdSelectSource_click()
    Dim FileName As String
    FileName = GetFile("Open: Merge Date Source File", 1)
    If FileName <> "" Then lblSource = LCase$(FileName)
End Sub

Private Sub Form_Load()
  If FolderOpen Then lblDest = LCase$(sBinderFileName)
End Sub


