VERSION 5.00
Begin VB.Form frmPrintDocument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Document"
   ClientHeight    =   2970
   ClientLeft      =   2355
   ClientTop       =   2700
   ClientWidth     =   5445
   ClipControls    =   0   'False
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
   ScaleHeight     =   2970
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
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
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
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
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
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
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
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
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
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
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame frmPrinter 
      Caption         =   "Printer Information"
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
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdFonts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Select"
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
         Left            =   4080
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Properties"
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
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Font: Default Printer Font"
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3855
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmPrintDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFonts_Click()
  Screen.MousePointer = 11
  Fonts.Show 1
  If SFont <> "" Then lblStatus = "Current Font: " & Trim(SFont)
End Sub



Private Sub Command1_Click()
DoPrint = True
Unload Me
End Sub

Private Sub Command2_Click()
DoPrint = False
Unload Me
End Sub

Private Sub Command3_Click()
x% = BDM_PrinterSetup()
GetPrinter
End Sub

Private Sub Form_Load()
nPrintMargin = 7
If SFont <> "" Then lblStatus = "Current Font: " & Trim(SFont)
GetPrinter
CenterMe Me
Screen.MousePointer = 0
End Sub

Private Sub GetPrinter()
 'Dim DefPrinter As typWindowsDevice
 'DUPRT_GetDefaultPrinter DefPrinter
 'lblPrinter = "Current Printer: " & DefPrinter.sWindowsDeviceUserName & " on " & DefPrinter.sWindowsDevicePortName
End Sub

Private Sub Option1_Click()
nPrintMargin = 0    'Do not print margins
End Sub

Private Sub Option2_Click()
nPrintMargin = 6
End Sub

Private Sub Option3_Click()
nPrintMargin = 7
End Sub

Private Sub Option4_Click()
nPrintMargin = 5
End Sub

