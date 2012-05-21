VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Binder Backup"
   ClientHeight    =   3780
   ClientLeft      =   1290
   ClientTop       =   2205
   ClientWidth     =   7680
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
   Icon            =   "BACKUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNewFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&New Folder"
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
      Left            =   6240
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
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
      Left            =   6240
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Backup Now"
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
      Left            =   6240
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2820
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Files Already in Folder:"
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
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblBackupTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Backup to:"
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
      Top             =   3480
      Width           =   7455
   End
   Begin VB.Label lblDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drives:"
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
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lblDirectorys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Folders:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cd As String
Dim Ph As String

Private Sub Backupnow(BPath As String)
 Dim i As Integer, BackFN  As Integer
 
 If right$(BPath, 1) <> "\" Then BPath = BPath & "\"
 BPath = BPath & Si
 On Error GoTo BackupError
 'Check destination for file
 If Dir$(BPath) <> "" Then  'File already exists
    i = MsgBox("The file " & UCase$(BPath) & crlf & "Already exists.  Replace it?", MB_ICONEXCLAIMATION + MB_YESNO, "File Exists")
    If i = IDNO Then Exit Sub
 End If
 Screen.MousePointer = 11
 BackFN = FreeFile
 Open BPath For Random As #BackFN Len = PD_Size
    For i = 1 To 102              'Copy Binder no need to decript
       Get #FreeNum, i, Pd
       Put #BackFN, i, Pd
    Next i
 Close #BackFN
 Screen.MousePointer = 0
 MsgBox "Backup complete", MB_ICONINFORMATION, "Backup Finished"
Exit Sub
BackupError:
  Beep
  MsgBox Error$(Err) & crlf & crlf & "Backup incomplete", MB_ICONSTOP, "Error"
Exit Sub
Resume
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdNewFolder_Click()
    Dim newDir As String, parentDir As String
    
    newDir = InputBox("Please enter the name for the new folder:")
    
    parentDir = Dir1.Path
    If right$(parentDir, 1) <> "\" Then parentDir = parentDir & "\"

    On Error Resume Next
    MkDir parentDir & newDir
    If Err Then
        MsgBox "An error has accord while attempting to create the new folder."
    End If
    Dir1.Refresh
End Sub

Private Sub Command1_Click()
 BPath$ = Dir1.Path
 Backupnow BPath$
 Si = ""
 Unload frmBackup
End Sub

Private Sub Command2_Click()
 Si = ""
 Unload frmBackup
End Sub

Private Sub Dir1_Change()
 lblBackupTo.Caption = "Backup To: " & Dir1.Path
 File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dim x As Integer, msg As String
  On Error GoTo DriveError
  Dir1.Path = Drive1.Drive
  Cd = Drive1.Drive
  ChDrive Cd         'Change the active drive
  Exit Sub

DriveError:
  If Err = 68 Then
      x = MsgBox("Cannot read from drive " & Drive1.Drive & crlf & "Make sure that the drive door is closed, and that the disk is formatted and free from errors", MB_ICONEXCLAMATION + MB_RETRYCANCEL, "Error Reading Drive " & Drive1.Drive)
      If x = IDRETRY Then
          Resume
      Else
          Drive1.Drive = Cd
          Dir1.Path = Cd
          Exit Sub
      End If
  Else
      msg = "The following unexpected error has accured while reading drive " & UCase(Drive1.Drive)
      If right(msg, 1) <> ":" Then
          msg = msg & ": "
      Else
          msg = msg & " "
      End If
      MsgBox msg & crlf & crlf & Error$, MB_ICONEXCLAMATION, "Unexpected Error"
      Drive1.Drive = Cd
      Exit Sub
  End If
Resume
End Sub

Private Sub Form_Load()
 Cd = Drive1.Drive
 lblBackupTo.Caption = "Backup To: " & Dir1.Path
End Sub


