VERSION 5.00
Begin VB.Form PrintBinderDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   5655
   ClientLeft      =   945
   ClientTop       =   1350
   ClientWidth     =   7650
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
   ScaleHeight     =   5655
   ScaleWidth      =   7650
   Begin VB.PictureBox picPrinting 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   6360
      ScaleHeight     =   1425
      ScaleWidth      =   5025
      TabIndex        =   33
      Top             =   3960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "  Printing"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label lblPrintStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Preparing to Print"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   4695
      End
      Begin VB.Image imgPrinter 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   120
         Picture         =   "PDIRPRIN.frx":0000
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblPrinting 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Printing binder"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Margin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   28
      Top             =   1680
      Width           =   2175
      Begin VB.OptionButton Option6 
         Caption         =   "1 Millimeter"
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
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "1 Centimeter"
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
         TabIndex        =   30
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "1 Inch"
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
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame frmPrinter 
      Caption         =   "Print Job Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdFonts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Page Setup"
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
         Left            =   3600
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Printer Properties"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblColumns 
         Caption         =   "Columns: 1"
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
         TabIndex        =   42
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblBodyProp 
         Caption         =   "Body: Times New Roman"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblHeadProp 
         Caption         =   "Header:  Arial"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblPrinter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Printer:"
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
         TabIndex        =   27
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox txtHeading 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "Healine"
         Top             =   480
         Width           =   3255
      End
      Begin VB.CheckBox chkHeading 
         Caption         =   "Page Header:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkBlank 
         Caption         =   "Don't Print Blank Fields"
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
         Left            =   2640
         TabIndex        =   22
         Top             =   840
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Notes"
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
         Index           =   10
         Left            =   240
         TabIndex        =   21
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Web Page"
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
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "E-Mail"
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
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Fax"
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
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Pager"
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
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Business Phone"
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
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Home Phone"
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
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Templates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2880
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
         Begin VB.OptionButton optAPO 
            Caption         =   "Default Template"
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
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton cmdResetTemplate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Reset Template"
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
            TabIndex        =   6
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optAPO 
            Caption         =   "Saved Template"
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
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "City       State       Zip"
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
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Address"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Name"
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkAddvancedOptions 
         Caption         =   "Company"
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
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6240
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Print"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdPrintBuffer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "V"
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
         Left            =   1920
         TabIndex        =   37
         Top             =   600
         Width           =   135
      End
      Begin VB.OptionButton optPrintBuffer 
         Caption         =   "Print Buffer"
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
         TabIndex        =   36
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectItems 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "V"
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
         Left            =   1920
         TabIndex        =   32
         Top             =   360
         Width           =   135
      End
      Begin VB.OptionButton optPrintBookmarkedItems 
         Caption         =   "Bookmarked Items"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optPrintEntireFolder 
         Caption         =   "Entire Binder"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optPrintSelectedItem 
         Caption         =   "Selected Item"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "PrintBinderDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DefaultTemplate As Integer
Dim TempFont As String

Dim i As Integer, Num As Integer, TotalItems As Integer
Dim X As Integer, Title As String, PrintTitle As Integer
Dim PrintBlankFields As Integer, PrintField(10) As Integer
Dim CurrentPrinter As String

Dim nColOffset As Integer, nCenterLine As Integer
Dim nMarginOffset As Integer, TopY As Double, nCurCol As Integer




Private Sub cmdFonts_Click()
    PrintStyles.Show 1
    ShowFontProp
End Sub



Private Sub cmdPrintBuffer_Click()
 Load frmPrintBuffer
 UpdatePrintBuffer
 frmPrintBuffer.Show 1

 If PrintBuffer(0) = 0 Then ' No Items
     cmdPrintBuffer.Enabled = False
     optPrintBuffer.Enabled = False
     optPrintSelectedItem.value = True
 Else
     optPrintBuffer.value = True
 End If
End Sub

Private Sub cmdResetTemplate_Click()
Dim i As Integer
For i = 0 To 10
   chkAddvancedOptions(i).value = 1
   SetMyIni "BDMAN.INI", "Print Template", "Field" & Trim$(Str$(i)), "1"
Next i
chkBlank.value = 1
chkHeading.value = 1
SetMyIni "BDMAN.INI", "Print Template", "PrintHeading", "1"
SetMyIni "BDMAN.INI", "Print Template", "PrintBlank", "1"
End Sub

Private Sub cmdSelectItems_Click()
 On Error Resume Next
 frmSelectItems.Show 1
 If UBound(SelectedItems) >= 1 Then
    optPrintSelectedItem.Caption = "Selected Items"
 Else
    optPrintSelectedItem.Caption = "Selected Item"
 End If
 optPrintSelectedItem.value = 1
End Sub

Private Sub Command1_Click()
    PrintNow
End Sub

Private Sub Command2_Click()
  Unload PrintBinderDialog
End Sub

Private Sub Command3_Click()
  X% = BDM_PrinterSetup()
  GetPrinter
End Sub

Private Sub Form_Load()
    Dim i As Integer, Flag As Integer, Fetch As String
 
    LoadPrintFormatting
 
    On Error Resume Next
    TempFont = SFont$
    Call optAPO_Click(0)
 
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
 

    If Si$ = "1" Then
        optPrintSelectedItem.value = True
    ElseIf Si$ = "2" Then
        optPrintSelectedItem.Enabled = False
        optPrintEntireFolder.value = True
    ElseIf Si = "3" Then
        optPrintBuffer.value = True
    End If


'---------------------------------------------------------------------
'Check to see if there are any bookmarks in Binder
'---------------------------------------------------------------------
    Flag = False
    For i = 2 To 100
        Get #FreeNum, i, Pd
        If Pd.Bookmark = True Then
            Flag = True
            Exit For
        End If
    Next i
    
    optPrintBookmarkedItems.Enabled = Flag 'Enable Bookmarked items option
    If PrintBuffer(0) = 0 Then             ' Disable print buffer
        optPrintBuffer.Enabled = False
        cmdPrintBuffer.Enabled = False
    End If
 
    GetPD FreeNum, CurrentIndex, True

    '--------------------------------------------------------
    'change advanced option button titles to user defined fields
    '----------------------------------------------------------
    chkAddvancedOptions(0).Caption = UserField(0)
    chkAddvancedOptions(1).Caption = UserField(1)
    chkAddvancedOptions(2).Caption = UserField(2)
    chkAddvancedOptions(3).Caption = UserField(3) & "   " & UserField(4) & "   " & UserField(5)
    chkAddvancedOptions(4).Caption = UserField(6)
    chkAddvancedOptions(5).Caption = UserField(7)
    chkAddvancedOptions(6).Caption = UserField(8)
    chkAddvancedOptions(7).Caption = UserField(9)
    chkAddvancedOptions(8).Caption = UserField(10)
    chkAddvancedOptions(9).Caption = UserField(11)
    chkAddvancedOptions(10).Caption = UserField(12)
    
    For i = 0 To 10
        If Trim$(chkAddvancedOptions(i).Caption) = "" Then
            chkAddvancedOptions(i).Visible = False
            chkAddvancedOptions(i).value = False
        End If
    Next i
    '----------------------------------------------------------
    'Fetch$ = GetMyIni("BDMAN.INI", "BDMAN", "FolderFont")
    'If Trim$(Fetch$) = "" Then
    '  SFont = Printer.FontName
    'Else
    '  SFont$ = Trim$(Fetch$)  'Get last used font from WIN.INI
    'End If
    'If SFont$ <> "" Then lblStatus = "Current Font: " & SFont$
     
    'Fetch = GetMyIni("BDMAN.INI", "BDMAN", "FolderColor")
    'nPrintColor = Val(Fetch)
    'lblStatus.ForeColor = QBColor(nPrintColor)
    'lblStatus.FontName = SFont
 
    ShowFontProp

    txtHeading = OpenFolderName
    GetPrinter
    nPrintMargin = 7
    Screen.MousePointer = 0 'Normal arrow

End Sub

Private Sub Form_Unload(Cancel As Integer)
 SetAsDefault = False
 SFont$ = TempFont
End Sub

Private Sub GetPrinter()
 'Dim DefPrinter As typWindowsDevice
 'DUPRT_GetDefaultPrinter DefPrinter
 'CurrentPrinter = DefPrinter.sWindowsDeviceUserName
 'lblPrinter = "Current Printer: " & CurrentPrinter & " on " & DefPrinter.sWindowsDevicePortName
End Sub

Private Function NumPrint() As Integer
 'Compute the number of Items that the user has selected
 'to be printed

 Dim i As Integer, NumItems As Integer
 
 NumItems = 2 'Allow for margins between Items
 For i = 0 To 9
    If PrintField(i) Then NumItems = NumItems + 1
 Next i
 If PrintField(10) Then NumItems = NumItems + 2  'Notes field can be long
 NumPrint = NumItems
End Function

Private Sub optAPO_Click(Index As Integer)
 Dim i As Integer, Fetch As String
 If Index = 0 Then
    DefaultTemplate = False
    For i = 0 To 10
      Fetch = GetMyIni("BDMAN.INI", "Print Template", "Field" & Trim$(Str$(i)))
      If Fetch <> "" Then chkAddvancedOptions(i).value = Val(Fetch)
    Next i
    Fetch = GetMyIni("BDMAN.INI", "Print Template", "PrintHeading")
    If Fetch <> "" Then chkHeading.value = Val(Fetch)
    Fetch = GetMyIni("BDMAN.INI", "Print Template", "PrintBlank")
    If Fetch <> "" Then chkBlank.value = Val(Fetch)
  Else
    DefaultTemplate = True
    For i = 0 To 10
       chkAddvancedOptions(i).value = 1
    Next i
    chkHeading.value = 1
    chkBlank.value = 1
 End If
End Sub

Private Sub Option4_Click()
nPrintMargin = 5
End Sub

Private Sub Option5_Click()
 nPrintMargin = 7
End Sub

Private Sub Option6_Click()
nPrintMargin = 6
End Sub

Private Sub PrintCurrentData()
'This procedure prints the data in the PD varible and
'Determins number of fields that will print and see if there is room on this
'page.
'
' Last Modified 3/11/2000
'----------------------------------------------------------------------------



' Set the Document Print Style and Begin Printing
SetBodyStyle

Dim NumItems As Integer, i   As Integer
Dim ItemHeight

ReDim PrintThisField(10) As Integer

For i = 0 To 10
   PrintThisField(i) = PrintField(i)
Next i

 '------------------------------------------------------------------------
 'Determin which fields wont be printed
 '------------------------------------------------------------------------
 
 NumItems = TotalItems
 If Not PrintBlankFields Then
      'Find which fields are blank so they won't be printed.
      If Trim(Pd.Company) = "" Then
          NumItems = NumItems - 1
          PrintThisField(0) = False
      End If
      If Trim(Pd.AName) = "" Then
          NumItems = NumItems - 1
          PrintThisField(1) = False
      End If
      If Trim(Pd.Address) = "" Then
          NumItems = NumItems - 1
          PrintThisField(2) = False
      End If
      If (Trim(Pd.City) = "") And (Trim(Pd.State) = "") And (Trim(Pd.Zip_Code) = "") Then
          NumItems = NumItems - 1
          PrintThisField(3) = False
      End If
      If Trim(Pd.Home_Phone) = "" Then
          NumItems = NumItems - 1
          PrintThisField(4) = False
      End If
      If Trim(Pd.Bus_Phone) = "" Then
          NumItems = NumItems - 1
          PrintThisField(5) = False
      End If
      If Trim(Pd.Pager) = "" Then
          NumItems = NumItems - 1
          PrintThisField(6) = False
      End If
      If Trim(Pd.Fax) = "" Then
          NumItems = NumItems - 1
          PrintThisField(7) = False
      End If
      If Trim(Pd.E_Mail) = "" Then
          NumItems = NumItems - 1
          PrintThisField(8) = False
      End If
      If Trim(Pd.WebPage) = "" Then
          NumItems = NumItems - 1
          PrintThisField(9) = False
      End If
      If Trim(Pd.Notes) = "" Then
          NumItems = NumItems - 2
          PrintThisField(10) = False
      End If
   End If
      
   '---------------------------------------------
   'Check to see if there is room on the paper
   '---------------------------------------------
   
   
    
   
        ' Estimate the Height in Twips of the to be printed item
        ItemHeight = Printer.TextHeight("ABCDEFGHIGKLMNOPQRSTUVWXYZ123409")
        ItemHeight = ItemHeight * (NumItems + 2)
   
        If Printer.ScaleHeight - (Printer.CurrentY + ItemHeight) < 0 Then
            ' If there are more than one column then
            ' the colum offset needs to be calculated
            ' and the cursor moved to the next column
            If BodyProp.nColumns > 1 And nCurCol = 1 Then
                ' Adjust the Active Column
                nColOffset = (Printer.ScaleWidth / 2)
                nCenterLine = 0
                linesPrinted = 0
                
                ' Set to Top of Column
                Printer.CurrentY = TopY
                nCurCol = 2
            Else
                Printer.NewPage
                nColOffset = 0
                nCenterLine = (Printer.ScaleWidth / 2)
                linesPrinted = 0
                nCurCol = 1
                AdjustPrintItHandlers 1, 0, nMarginOffset, 0
                PrintPageHeader
            End If
        End If
   
    
      
      
SetBodyStyle

'--------------------------------------------------------------------------
'   Print Record
'   After printing the first line, find the Y coord of the top line
'--------------------------------------------------------------------------

If PrintThisField(0) Then
    AddMargin nPrintMargin
    'Printer.Print
    PrintLine UserField(0) + ": " + Trim$(Pd.Company), nColOffset, nCenterLine
End If

If PrintThisField(1) Then
    AddMargin nPrintMargin
    PrintLine UserField(1) + ": " + Trim$(Pd.AName), nColOffset, nCenterLine
End If


If PrintThisField(2) Then
    AddMargin nPrintMargin
    PrintLine UserField(2) + ": " + Trim$(Pd.Address), nColOffset, nCenterLine
End If

If PrintThisField(3) Then
  AddMargin nPrintMargin
  PrintLine UserField(3) + ": " + RTrim$(Pd.City) + "   " + UserField(4) + ": " + RTrim$(Pd.State) + "   " + UserField(5) + ": " + RTrim$(Pd.Zip_Code), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(4) Then
    PrintLine UserField(6) + ": " + Trim$(Pd.Home_Phone), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(5) Then
    PrintLine UserField(7) + ": " + Trim$(Pd.Bus_Phone), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(6) Then
    PrintLine UserField(8) + ": " + Trim$(Pd.Pager), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(7) Then
    PrintLine UserField(9) + ": " + Trim$(Pd.Fax), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(8) Then
    PrintLine UserField(10) + ": " + Trim$(Pd.E_Mail), nColOffset, nCenterLine
End If

AddMargin nPrintMargin
If PrintThisField(9) Then
    PrintLine UserField(11) + ": " + Trim$(Pd.WebPage), nColOffset, nCenterLine
End If

'Don't count notes field "PrintLine" will count
'the lines it takes to print the notes
linesPrinted = linesPrinted + NumItems - 2

'Notes field may need Word Wraping
If PrintThisField(10) Then
    PrintLine UserField(12) & ": " & Trim$(Pd.Notes), nColOffset, nCenterLine
End If

Printer.Print

'----------------------------------------------------------------------------
End Sub

Private Sub PrintNow()
'Print the information in a Binder
Dim i As Integer
Dim msg As String
'---------------------------------------------------------------
'Prepare to Start Printing
'---------------------------------------------------------------

SetBodyStyle

Screen.MousePointer = 11

lblPrinting = "Printing " & frmMain.lblFolder
picPrinting.Top = 1680
picPrinting.Left = 1440
picPrinting.Visible = True
DoEvents

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

cmdFonts.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False
DoEvents

'Set Page Header
Title = txtHeading.text
Title = RTrim$(Title)
PrintTitle = chkHeading.value
PrintBlankFields = Not chkBlank.value

'Set Print Template in an Array to save time later.
For i = 0 To 10
   PrintField(i) = chkAddvancedOptions(i).value
Next i

'Save the print template if the deafault template setting is not selected.
If Not DefaultTemplate Then SavePrintTemplate

'--------------------------------------------------------------------------

TotalItems = NumPrint() 'Find number of Items the user wants to print

PrintPageHeader    'Print the page header

' Set Initial Print Tracking Values
' Added 3/10/2000

' If multi columns are in use, find the Center Line
If BodyProp.nColumns > 1 Then
    nCenterLine = (Printer.ScaleWidth / 2)
Else
    nCenterLine = 0
End If

' Find the Margin Offset
AddMargin nPrintMargin
nMarginOffset = Printer.CurrentX

nColOffset = 0
TopY = 0
TopY = Printer.CurrentY
nCurCol = 1

' End Set Initial....

If optPrintSelectedItem.value Then
    For i = 0 To UBound(SelectedItems)
      GetPD FreeNum, SelectedItems(i), True
      PrintCurrentData
    Next i
ElseIf optPrintBuffer.value Then
    For i = 0 To UBound(PrintBuffer)
      GetPD FreeNum, PrintBuffer(i), True
      PrintCurrentData
    Next i
ElseIf optPrintBookmarkedItems.value Then
     For i = 2 To 100
          Get #FreeNum, i, Pd
          If Pd.Bookmark = True Then
              GetPD FreeNum, i, True
              PrintCurrentData
          End If
     Next i
ElseIf optPrintEntireFolder.value Then
     For i = 2 To 100
        GetPD FreeNum, i, True
        If Trim(Pd.Company) = "" And Trim(Pd.AName) = "" Then Exit For
        PrintCurrentData
     Next i
Else
     'This should never happen
     MsgBox "There is no print option selected", MB_ICONSTOP, "Error"
     Unload Me
     Exit Sub
End If
Printer.EndDoc
Unload Me
Exit Sub

'---------------------------------------------------------------------------
ErrorPrint:
  Screen.MousePointer = 0
  Beep
  MsgBox "An unexpected error has occurred while printing.", MB_ICONSTOP, "Print Error"
  Printer.EndDoc
  Unload Me
  Exit Sub
Resume

End Sub

Private Sub PrintPageHeader()
  Dim pge As Integer
  
  Printer.Print
  Printer.Print
  Printer.Print
  
  linesPrinted = linesPrinted + 3
  
  SetHeaderStyle
  
  If PrintTitle = 1 Then
      
      AddMargin nPrintMargin
      
      Printer.Print Title;
      
      If HeadingProp.bPageNumbers Then Printer.Print " - Page "; Printer.Page;
      
      Printer.Print Space$(160)
      
      Printer.Print
      linesPrinted = linesPrinted + 2
      
  End If
  pge = Printer.Page
  lblPrintStatus = "Printing page " & Trim$(Str$(pge))
  
  SetBodyStyle
  DoEvents
End Sub

Private Sub SavePrintTemplate()
Dim temp As Integer
For i = 0 To 10
   SetMyIni "BDMAN.INI", "Print Template", "Field" & Trim$(Str$(i)), Trim$(Str$(PrintField(i)))
Next i
temp = chkHeading.value
SetMyIni "BDMAN.INI", "Print Template", "PrintHeading", Trim$(Str$(temp))
temp = chkBlank.value
SetMyIni "BDMAN.INI", "Print Template", "PrintBlank", Trim$(Str$(temp))
End Sub

' These two functions Style functions were added on 3/10/2000

Public Sub SetBodyStyle()
    Printer.Font.Name = BodyProp.sName
    Printer.Font.size = BodyProp.nSize
    Printer.Font.Bold = BodyProp.bBold
    Printer.Font.Italic = BodyProp.bItalic
    Printer.Font.Underline = BodyProp.bUnderline
    Printer.FontStrikethru = BodyProp.bStrikethru
    Printer.ForeColor = BodyProp.nColor
End Sub

Public Sub SetHeaderStyle()
    Printer.Font.Name = HeadingProp.sName
    Printer.Font.size = HeadingProp.nSize
    Printer.Font.Bold = HeadingProp.bBold
    Printer.Font.Italic = HeadingProp.bItalic
    Printer.Font.Underline = HeadingProp.bUnderline
    Printer.FontStrikethru = HeadingProp.bStrikethru
    Printer.ForeColor = HeadingProp.nColor
End Sub

Public Sub ShowFontProp()
'  Show the Printer and Page Information
'  Last Modified 3/11/2000

    lblHeadProp = "Header: " + HeadingProp.sName + " (" + Format(HeadingProp.nSize) + "pt.)"
    lblHeadProp.ForeColor = HeadingProp.nColor
    lblHeadProp.Font.Bold = HeadingProp.bBold
    lblHeadProp.Font.Italic = HeadingProp.bItalic
    lblHeadProp.Font.Underline = HeadingProp.bUnderline
    lblHeadProp.FontStrikethru = HeadingProp.bStrikethru
    
    lblBodyProp = "Body: " + BodyProp.sName + " (" + Format(BodyProp.nSize) + "pt.)"
    lblBodyProp.ForeColor = BodyProp.nColor
    lblBodyProp.Font.Bold = BodyProp.bBold
    lblBodyProp.Font.Italic = BodyProp.bItalic
    lblBodyProp.Font.Underline = BodyProp.bUnderline
    lblBodyProp.FontStrikethru = BodyProp.bStrikethru
    
    lblColumns = "Columns: " + Format(BodyProp.nColumns)
    lblColumns = lblColumns + "    Page Numbers: "
    
    If HeadingProp.bPageNumbers Then
        lblColumns = lblColumns + "Yes"
    Else
        lblColumns = lblColumns + "No"
    End If
   
End Sub
