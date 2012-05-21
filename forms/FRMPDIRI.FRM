VERSION 5.00
Begin VB.Form ViewWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Viewer"
   ClientHeight    =   5235
   ClientLeft      =   150
   ClientTop       =   1305
   ClientWidth     =   9315
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FRMPDIRI.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNextItem 
      Caption         =   ">>>"
      Height          =   375
      Left            =   720
      TabIndex        =   52
      ToolTipText     =   "Go to the next item in the binder"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdPreviousItem 
      Caption         =   "<<<"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      ToolTipText     =   "Go to the previous item in the binder"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Help"
      Height          =   375
      Left            =   8280
      TabIndex        =   50
      ToolTipText     =   "Explain how this dialog is used."
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Print"
      Height          =   375
      Left            =   8280
      TabIndex        =   49
      ToolTipText     =   "Print the Current Item"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdBookmark 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Remove &Bookmark"
      Height          =   375
      Left            =   5160
      TabIndex        =   48
      ToolTipText     =   "Add or Remove the bookmark from this item"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Letter"
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   47
      ToolTipText     =   "Create a letter based on this item"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Memo"
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   46
      ToolTipText     =   "Create a Memo based on this item"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox chkCopyAllFields 
      Caption         =   "Copy All Fields"
      Height          =   375
      Left            =   1560
      TabIndex        =   45
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   44
      ToolTipText     =   "Close this window"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Delete"
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      ToolTipText     =   "Delete this item"
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkCopyLabels 
      Caption         =   "Copy Field Labels"
      Height          =   375
      Left            =   3240
      TabIndex        =   39
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdModify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Modify"
      Height          =   375
      Left            =   8280
      TabIndex        =   38
      ToolTipText     =   "Open this item in the edit window"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   7
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCopyToClip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Send to Clipboard"
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Image imgWebLink 
      Height          =   165
      Left            =   7800
      Picture         =   "FRMPDIRI.frx":030A
      ToolTipText     =   "Goto Web Site"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Image imgEmailLink 
      Height          =   165
      Left            =   7800
      Picture         =   "FRMPDIRI.frx":0588
      ToolTipText     =   "Send E-Mail"
      Top             =   3360
      Width           =   255
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   9360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   8160
      X2              =   9360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   8160
      X2              =   8160
      Y1              =   480
      Y2              =   4680
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   2040
      TabIndex        =   43
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   1470
      TabIndex        =   42
      Top             =   4080
      Width           =   465
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Use the Check Box Next to Each Field to Select the Ones you Wish to Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   12
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page (URL):"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   660
      TabIndex        =   13
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   14
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   1470
      TabIndex        =   15
      Top             =   3360
      Width           =   465
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9360
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   16
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   18
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   19
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   21
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   23
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   24
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   25
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   26
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   1635
      TabIndex        =   27
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pager:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   1470
      TabIndex        =   28
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bus. Phone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   1065
      TabIndex        =   29
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   30
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      Height          =   195
      Index           =   5
      Left            =   6345
      TabIndex        =   31
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   195
      Index           =   4
      Left            =   4155
      TabIndex        =   32
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   33
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   35
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   34
      Top             =   480
      Width           =   705
   End
End
Attribute VB_Name = "ViewWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status As Integer
Dim DoNotShow As Integer
Dim bttn As Double
Dim bttnPush As Integer

Private Sub chkCopyAllFields_Click()
Dim i As Integer
If chkCopyAllFields.value = 1 Then
   For i = 0 To 12
      ChkCopy(i).value = 1
   Next i
Else
   LoadTemplate
End If

End Sub

Private Sub cmdBookmark_Click()
 GetPD FreeNum, CurrentIndex, True
 If Pd.Bookmark = True Then
    Pd.Bookmark = False
    cmdBookmark.Caption = "Add Bookmark"
 Else
    Pd.Bookmark = True
    cmdBookmark.Caption = "Remove Bookmark"
 End If
 PutPD FreeNum, CurrentIndex, True
 LoadBookmarks
End Sub

Private Sub cmdClose_Click()
 Dim i As Integer
 If Status = 0 Then
    Unload ViewWindow
 ElseIf Status = 1 Then
    For i = 0 To 12
       ChkCopy(i).Visible = False
    Next i
    chkCopyLabels.Visible = False
    chkCopyAllFields.Visible = False
    cmdModify.Visible = True
    cmdPrint.Visible = True
    cmdDelete.Visible = True
    cmdBookmark.Visible = True
    
    Status = 0
    cmdClose.Caption = "&Close"
    cmdCopyToClip.Caption = "&Send to Clipboard"
    cmdMemo(0).Visible = True
    cmdMemo(1).Visible = True
    
    lblSelect.Visible = False
  End If
End Sub

Private Sub cmdCopyToClip_Click()
' Allow the user to copy information to the clipboard.

Dim i As Integer, CTCB As String, temp As String, DataFlag As Integer
Dim BDM20 As FolderType

Screen.MousePointer = 11
Get #FreeNum, 1, BDM20
If Status = 0 Then
    '--------------------------
    ' Hide the option buttons
    '--------------------------
    cmdModify.Visible = False
    cmdPrint.Visible = False
    cmdDelete.Visible = False
    cmdBookmark.Visible = False
    cmdMemo(0).Visible = False
    cmdMemo(1).Visible = False
    
    '----------------------------------------

    LoadTemplate    ' Recheck the items the user last checked

    chkCopyLabels.value = 0
    chkCopyLabels.Visible = True
    chkCopyAllFields.value = 0
    chkCopyAllFields.Visible = True
    
    lblSelect.Visible = True
    Status = 1
    cmdClose.Caption = "Cancel"
    cmdCopyToClip.Caption = "Copy Now"
ElseIf Status = 1 Then
    GoSub CopyNow
    lblSelect.Visible = False
    t$ = ""
    For i = 0 To 12
       vl% = ChkCopy(i).value
       t$ = t$ & Trim$(Str$(vl%))
    Next i
    BDM20.Fax = t$
    Put #FreeNum, 1, BDM20
    Call cmdClose_Click
End If
    Me.Refresh
Screen.MousePointer = 0
Exit Sub

CopyNow:
  CTCB = ""
  'Add checked items to string
  'Note:  Remember that a varible followed by a % is of the Integer data type
  CopyLabels% = chkCopyLabels     'Load checked state into a Integer varible
                                  'for faster access in the following loop.
  For i = 0 To 12
    If ChkCopy(i).value = 1 Then
        temp = ""
        
        'If copylabels is true then copy the field labels as well as
        'the fields.
        If CopyLabels% = 1 Then temp = temp & Label1(i).Caption & " "
        temp = temp & Trim$(lblDisplay(i))  'Add Item to end of string
           
        If Trim$(CTCB) <> "" Then CTCB = CTCB & crlf  ' Start a new line
        CTCB = CTCB & temp
    End If
  Next i
  
  'Copy the string to Windows Clipboard
  Clipboard.SetText CTCB

Return
End Sub

Private Sub cmdDelete_Click()
 Unload ViewWindow
 DeleteItem
End Sub

Private Sub cmdHelp_Click()
    NoHelp
End Sub

Private Sub cmdMemo_Click(Index As Integer)
 Screen.MousePointer = 11
 Dim Letter As String, BR As Integer
 BR = NoteBook
 
 If Index = 0 Then
    Letter = "Date: " & Format$(Date$, "mmmm d, yyyy") & crlf
    Letter = Letter & "To: " & Trim$(Pd.AName) & crlf
    Letter = Letter & "From:" & crlf
    Letter = Letter & "Subject:" & crlf
 ElseIf Index = 1 Then
    Load LetterWizard
    LetterWizard.txtAddresse = Trim$(lblDisplay(0)) & crlf & Trim$(lblDisplay(2)) & crlf & Trim$(lblDisplay(3))
    If Trim$(lblDisplay(3)) <> "" Then LetterWizard.txtAddresse = LetterWizard.txtAddresse & ",  " & Trim$(lblDisplay(4)) & "  " & Trim$(lblDisplay(5))
    LetterWizard.txtSendTo = Trim$(lblDisplay(1))
    Screen.MousePointer = 0
    LetterWizard.Show 1
    Letter = SelFormat.SFormat
 End If
 
 If Letter <> "" Then
    frmMDI.Show
    Screen.MousePointer = 11
    If BR = 1 Then FileNew
    If frmMDI.WindowState = 1 Then frmMDI.WindowState = 0
    frmMDI.ActiveForm.Text1 = Letter
    GetRecentFiles
 End If
 Screen.MousePointer = 0
End Sub

Private Sub cmdModify_Click()
    Unload ViewWindow
    EditItem
    ShowRecord CurrentIndex
End Sub

Private Sub cmdNextItem_Click()
    ' Load the next item into view
    If CurrentIndex - 1 < ItemCount Then
        ShowUnselectedRow frmMain.flexItems, CurrentIndex - 1
        CurrentIndex = CurrentIndex + 1
        ShowRecord CurrentIndex
        ShowSelectedRow frmMain.flexItems, CurrentIndex - 1
    End If
End Sub

Private Sub cmdPreviousItem_Click()
    ' Load the Previous item into view
    If CurrentIndex >= 2 Then
        ShowUnselectedRow frmMain.flexItems, CurrentIndex - 1
        CurrentIndex = CurrentIndex - 1
        ShowRecord CurrentIndex
        ShowSelectedRow frmMain.flexItems, CurrentIndex - 1
    End If
End Sub

Private Sub cmdPrint_Click()
 Unload ViewWindow
 GetPD FreeNum, CurrentIndex, True
 Si = "1"
 PrintBinderDialog.Show 1
End Sub



Private Sub Form_Load()
   CenterMe Me
   Status = 0
   ViewWindowOpen = True
   For i% = 0 To 12
      If Trim$(UserField(i%)) <> "" Then
          Label1(i%) = UserField(i%) & ":"
      Else
          Label1(i%).Visible = False
          lblDisplay(i%).Visible = False
      End If
   Next i%
     
   If lblDisplay(10).Visible = False Then imgEmailLink.Visible = False
   If lblDisplay(11).Visible = False Then imgWebLink.Visible = False
   'frmMain.lstNames.Enabled = False
   
   If Si <> "&H000005&" Then
      ShowRecord CurrentIndex
      DoNotShow = False
   Else
      DoNotShow = True
   End If

   Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 0 Then
        bttn = 0
        Cls
    End If
End Sub

Private Sub Form_Paint()
 Dim i As Integer
 LineIn3D Line1
 LineIn3D Line2
 LineIn3D Line3
 LineIn3D Line4
 
 If lblSelect.Visible Then ControlIn3D lblSelect, 1, 1, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ViewWindowOpen = False
    'frmMain.lstNames.Enabled = True
End Sub

Private Sub LoadTemplate()
    ' The options the user last checked are stored in the
    ' fax field of the header item.

    Dim BDM20 As FolderType, temp As String, i As Integer
    
    Get #FreeNum, 1, BDM20
    temp = BDM20.Fax
    For i = 0 To 12
      If Trim$(UserField(i)) <> "" Then
        ChkCopy(i).value = Val(Mid$(temp, i + 1, 1))
        ChkCopy(i).Visible = True
      Else
        ChkCopy(i).value = 0   ' Disable Check box
      End If
    Next i
End Sub

Private Sub imgEmailLink_Click()
    Dim w As String
    w = Trim$(lblDisplay(10))
    If w <> "" Then
        OpenURL "mailto:" & w
        Cls
    End If
End Sub

Private Sub imgEmailLink_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgEmailLink_MouseMove button, Shift, x, y
End Sub

Private Sub imgEmailLink_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.1 Then
        Erase3DBorder imgWebLink, 1, 0, False
        bttnPush = False
        bttn = 1.1
    End If
    ControlIn3D imgEmailLink, 1, 0, bttnPush
End Sub

Private Sub imgEmailLink_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgEmailLink_MouseMove button, Shift, x, y
End Sub

Private Sub imgWebLink_Click()
    Dim w As String
    w = Trim$(lblDisplay(11))
    If w <> "" Then
        OpenURL w
        Cls
    End If
End Sub

Private Sub imgWebLink_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = True
    imgWebLink_MouseMove button, Shift, x, y
End Sub

Private Sub imgWebLink_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If bttn <> 1.2 Then
        Erase3DBorder imgEmailLink, 1, 0, False
        bttnPush = False
        bttn = 1.2
    End If
    ControlIn3D imgWebLink, 1, 0, bttnPush
End Sub

Private Sub imgWebLink_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    bttnPush = False
    imgWebLink_MouseMove button, Shift, x, y
End Sub



