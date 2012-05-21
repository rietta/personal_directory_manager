VERSION 5.00
Begin VB.Form frmFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Field Template Manager"
   ClientHeight    =   5520
   ClientLeft      =   915
   ClientTop       =   1530
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FRMFIELD.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   3120
      TabIndex        =   25
      Top             =   480
      Width           =   3735
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "Company"
         Text            =   "Company"
         ToolTipText     =   "The default title for this field is Company"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   0
         Left            =   2625
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "Name"
         Text            =   "Name"
         ToolTipText     =   "The default title for this field is Name"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   1
         Left            =   2625
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   2
         Left            =   825
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Address"
         Text            =   "Address"
         ToolTipText     =   "The default title for this field is Address"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   2
         Left            =   2625
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   3
         Left            =   840
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "City"
         Text            =   "City"
         ToolTipText     =   "The default title for this field is City"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   3
         Left            =   2625
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   4
         Left            =   825
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "State"
         Text            =   "State"
         ToolTipText     =   "The default title for this field is State"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   4
         Left            =   2625
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   5
         Left            =   825
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Zip"
         Text            =   "Zip"
         ToolTipText     =   "The default title for this field is Zip Code"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   5
         Left            =   2625
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   6
         Left            =   840
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "Home Phone"
         Text            =   "Home Phone"
         ToolTipText     =   "The default title for this field is Home Phone"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   6
         Left            =   2625
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   7
         Left            =   825
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "Bus. Phone"
         Text            =   "Bus. Phone"
         ToolTipText     =   "The default title for this field is Business Phone"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   7
         Left            =   2625
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   8
         Left            =   825
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Pager"
         Text            =   "Pager"
         ToolTipText     =   "The default title for this field is Pager"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   8
         Left            =   2625
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   9
         Left            =   825
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "Fax"
         Text            =   "Fax"
         ToolTipText     =   "The default title for this field is Fax"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   9
         Left            =   2625
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   10
         Left            =   825
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "E-Mail"
         Text            =   "E-Mail"
         ToolTipText     =   "The default title for this field is E-Mail"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   10
         Left            =   2625
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   11
         Left            =   840
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Web Page (URL)"
         Text            =   "Web Page (URL)"
         ToolTipText     =   "The default title for this field is Website"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   11
         Left            =   2625
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtFieldNames 
         Height          =   285
         Index           =   12
         Left            =   825
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Notes"
         Text            =   "Notes"
         ToolTipText     =   "The default title for this field is Notes"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton cdmReset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reset"
         Height          =   255
         Index           =   12
         Left            =   2625
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 1:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   100
         TabIndex        =   52
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 2:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   100
         TabIndex        =   51
         Top             =   600
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 3:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   100
         TabIndex        =   50
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 4:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   100
         TabIndex        =   49
         Top             =   1320
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 5:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   100
         TabIndex        =   48
         Top             =   1680
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 6:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   100
         TabIndex        =   47
         Top             =   2040
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 7:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   100
         TabIndex        =   46
         Top             =   2400
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 8:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   100
         TabIndex        =   45
         Top             =   2760
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 9:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   100
         TabIndex        =   44
         Top             =   3120
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 10:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   100
         TabIndex        =   43
         Top             =   3480
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 11:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   100
         TabIndex        =   42
         Top             =   3840
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 12:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   100
         TabIndex        =   41
         Top             =   4200
         Width           =   2000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Field 13:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   100
         TabIndex        =   40
         Top             =   4560
         Width           =   2000
      End
   End
   Begin VB.Frame famTemplates 
      Caption         =   "Templates"
      Height          =   3975
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   2895
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1560
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTemplateName 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdRename 
         Caption         =   "&Rename"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Delete the Selected Template"
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         ToolTipText     =   "Create a New Template"
         Top             =   3480
         Width           =   615
      End
      Begin VB.ListBox lstTemplates 
         Height          =   3180
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.PictureBox picTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   220
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   3105
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "The Default for This Field is Company"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2970
      End
   End
   Begin VB.CheckBox chkTip 
      Caption         =   "Show default title for field in yellow popup box while editing."
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Default"
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Save"
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblInstructions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Definable Fields Allow You to Change the Title Descriptions of the Fields in a Binder"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TemplateEditInfo(100, 12) As String
Dim TemplateEditIndex() As String
Dim CurrentEditTemplate As Integer
Dim EditIndex As Integer

Private Sub cdmReset_Click(Index As Integer)
    txtFieldNames(Index) = txtFieldNames(Index).Tag
End Sub

Private Sub cdmReset_GotFocus(Index As Integer)
    picTip.Visible = False
End Sub

Private Sub chkTip_Click()
    picTip.Visible = chkTip.value
End Sub

Private Sub cmdAdd_Click()
    If cmdAdd.Caption = "&OK" Then
        If EditIndex > UBound(TemplateEditIndex) Then
            ReDim Preserve TemplateEditIndex(EditIndex) As String
        End If
        TemplateEditIndex(EditIndex) = txtTemplateName
        If FolderOpen Then
            RefreshList EditIndex + 1
        Else
            RefreshList EditIndex
        End If
        ShowTemplateEditField False
    Else
        ' Create a New Index
        EditIndex = lstTemplates.ListCount - 1
        If Not FolderOpen Then EditIndex = EditIndex + 1
        txtTemplateName = ""
        ShowTemplateEditField True
    End If
End Sub

Private Sub cmdCancel_Click()
    If txtTemplateName.Visible Then
        ShowTemplateEditField False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCancel2_Click()
    cmdCancel_Click
End Sub

' Delete Template
' Created 1/22/2000
Private Sub cmdDelete_Click()
    Dim i As Integer, j As Integer
    
    EditIndex = CurrentEditTemplate - 1
    TemplateEditIndex(EditIndex) = ""
    
    ' Copy to an Intermediate Array
    ReDim DelList(UBound(TemplateEditIndex)) As String
    
    For i = 0 To UBound(TemplateEditIndex)
        DelList(i) = TemplateEditIndex(i)
    Next i
    
    If UBound(TemplateEditIndex) >= 1 Then
        ReDim TemplateEditIndex(UBound(TemplateEditIndex) - 1)
    Else
        ReDim TemplateEditIndex(0) As String
    End If
    
    j = 0
    For i = 0 To UBound(DelList)
        If Trim(DelList(i)) <> "" Then
            TemplateEditIndex(j) = DelList(i)
            j = j + 1
        End If
    Next i
    
    
    RefreshList EditIndex - 1
End Sub



Private Sub cmdRename_Click()
    EditIndex = CurrentEditTemplate - 1
    txtTemplateName = TemplateEditIndex(EditIndex)
    ShowTemplateEditField True
End Sub

Private Sub Command1_Click()
    Dim User1 As String, User2 As String, i As Integer
    Screen.MousePointer = 11
    If Trim$(txtFieldNames(0)) = "" Then
        MsgBox "Field 1 cannot be blank", MB_ICONEXCLAMATION, "Blank Field Detected"
        Exit Sub
    End If
 
    ResetPD
    
    SaveTemplateData TemplateEditIndex(), TemplateEditInfo()
    
    If lstTemplates.text = "[Current Binder]" Then 'Save field names in the 102 record after dump
            For i = 0 To 12
                UserField(i) = txtFieldNames(i)
            Next i
            SaveFields
    'ElseIf optAllFolders Then
    '    For i = 0 To 12
    '        temp$ = txtFieldNames(i)
    '        SetMyIni "BDMAN.INI", "Fields", "Field" & Trim$(Str$(i)), temp$
    '    Next i
    End If
    
    MsgBox "Your changes have been saved", MB_ICONINFORMATION, SPROGRAMNAME
    ResetPD
    Screen.MousePointer = 0
    
    InitializeGrid frmMain.flexItems
    Unload frmFields
End Sub

Private Sub Command2_Click()
 Dim i As Integer
 For i = 0 To 12
    txtFieldNames(i) = txtFieldNames(i).Tag
 Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass

    lstTemplates.Clear
    If FolderOpen Then lstTemplates.AddItem "[Current Binder]"
    InitializeTemplateInfo
    On Error Resume Next
    ReDim TemplateEditIndex(UBound(TemplateIndex))
    If Err = 9 Then
        MsgBox "Could not load data from template file", vbCritical, SPROGRAMNAME
    Else
        For i = 0 To UBound(TemplateIndex)
        lstTemplates.AddItem TemplateIndex(i)
        lstTemplates.ItemData(lstTemplates.NewIndex) = i
        TemplateEditIndex(i) = TemplateIndex(i)
        Next i
    
    End If
    ' If a Binder is open then load its
    ' internal fields
    LoadAllTemplateData TemplateEditInfo()
    If lstTemplates.ListCount >= 1 Then
        lstTemplates.ListIndex = 0
        
        LoadFields
        For i = 0 To 12
            If UserField(i) <> "" Then
                TemplateEditInfo(0, i) = UserField(i%)
            Else
                TemplateEditInfo(0, i) = ""
            End If
        Next i
    End If
    
    SwitchTemplate
    Screen.MousePointer = vbNormal
End Sub

Private Sub LoadIniFields()
       Dim temp As String
       Dim i As Integer
       For i = 0 To 12
         temp$ = GetMyIni("BDMAN.INI", "Fields", "Field" & Trim$(Str$(i)))
         If temp$ <> "" Then
            txtFieldNames(i) = temp$
         Else
            txtFieldNames(i) = txtFieldNames(i).Tag
         End If
       Next i
End Sub

Private Sub lstTemplates_Click()
    ' The first index (0) in the template data array is reserved for the
    ' current binder's internal field definitions
    If FolderOpen Then
        CurrentEditTemplate = lstTemplates.ListIndex
    Else
        CurrentEditTemplate = lstTemplates.ListIndex + 1
    End If
    If FolderOpen And CurrentEditTemplate = 0 Then
        cmdRename.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdRename.Enabled = True
        cmdDelete.Enabled = True
    End If
    
    SwitchTemplate
End Sub

Private Sub lstTemplates_GotFocus()
    picTip.Visible = False
End Sub

Private Sub txtFieldNames_Change(Index As Integer)
    ' Save TemplateInfo
    TemplateEditInfo(CurrentEditTemplate, Index) = txtFieldNames(Index).text
End Sub

Private Sub txtFieldNames_GotFocus(Index As Integer)
    
    'picTip.Top = txtFieldNames(Index).Top + txtFieldNames(Index).Height + 20
    lblTip = "The default title for this field is " & txtFieldNames(Index).Tag
    
    'picTip.Left = txtFieldNames(Index).Left + 50

    txtFieldNames(Index).SelStart = 0
    txtFieldNames(Index).SelLength = Len(txtFieldNames(Index))
    picTip.Visible = chkTip.value
End Sub

Private Sub txtFieldNames_KeyPress(Index As Integer, KeyAscii As Integer)
 If Chr$(KeyAscii) = ":" Then KeyAscii = 0
 KeyAscii = CheckChar(KeyAscii)  'Check for valid charactors
End Sub


Public Sub SwitchTemplate()
    Dim i As Integer
    For i = 0 To 12
        txtFieldNames(i).text = TemplateEditInfo(CurrentEditTemplate, i)
    Next i
End Sub

Public Sub ShowTemplateEditField(nShow As Boolean)
    txtTemplateName.Visible = nShow
    cmdDelete.Visible = Not nShow
    cmdRename.Visible = Not nShow
    cmdCancel2.Visible = nShow
    If nShow Then
        cmdAdd.Caption = "&OK"
        txtTemplateName.SetFocus
    Else
        cmdAdd.Caption = "&Add"
    End If
End Sub



Private Sub txtTemplateName_GotFocus()
    picTip.Visible = False
    txtTemplateName.SelStart = 0
    txtTemplateName.SelLength = Len(txtTemplateName)
End Sub
' Show the Templates in the List
Public Sub RefreshList(LstIndex As Integer)
    Dim i As Integer
    lstTemplates.Clear
    
    If FolderOpen Then lstTemplates.AddItem "[Current Binder]"

    For i = 0 To UBound(TemplateEditIndex)
        lstTemplates.AddItem TemplateEditIndex(i)
    Next i
    
    If LstIndex <= lstTemplates.ListCount - 1 Then
        lstTemplates.ListIndex = LstIndex
    Else
        lstTemplates.ListIndex = 0
    End If
End Sub
