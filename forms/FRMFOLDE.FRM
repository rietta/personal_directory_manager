VERSION 5.00
Begin VB.Form BinderDescriptionDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Binder"
   ClientHeight    =   1695
   ClientLeft      =   1635
   ClientTop       =   3645
   ClientWidth     =   7065
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   7065
   Begin VB.ComboBox cboTemplates 
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
      ItemData        =   "FRMFOLDE.frx":0000
      Left            =   1080
      List            =   "FRMFOLDE.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Continue"
      Default         =   -1  'True
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
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtFolderName 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Binder descriptions are not the same as the file name of the binder."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label lblTemplate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Template:"
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "BinderDescriptionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 Si = ""
 Unload BinderDescriptionDialog
End Sub

Private Sub cmdHelp_Click()
NoHelp
End Sub

Private Sub cmdOK_Click()
    Si = txtFolderName.text
    Si = Trim$(Si)
    If cboTemplates.Visible Then CurrentFolderTemplate = cboTemplates.ListIndex
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    cboTemplates.Clear
    On Error Resume Next
    For i = 0 To UBound(TemplateIndex)
        cboTemplates.AddItem TemplateIndex(i)
    Next i
    cboTemplates.ListIndex = 0
    CenterMe Me
End Sub



Private Sub txtFolderName_Change()
 t$ = txtFolderName
 If Trim$(t$) <> "" Then cmdOK.Enabled = True Else cmdOK.Enabled = False
End Sub

Private Sub txtFolderName_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckChar(KeyAscii)  'Check for valid charactors
End Sub

