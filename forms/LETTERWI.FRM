VERSION 5.00
Begin VB.Form LetterWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Letter"
   ClientHeight    =   5145
   ClientLeft      =   210
   ClientTop       =   1410
   ClientWidth     =   9330
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
   ScaleHeight     =   5145
   ScaleWidth      =   9330
   Begin VB.TextBox txtSendTo 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Frame famPlaceMyAddress 
      Caption         =   "Place My Address"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
      Begin VB.OptionButton optPlaceMyAddress 
         Caption         =   "At the bottom of the letter"
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
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optPlaceMyAddress 
         Caption         =   "At the top of the letter"
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
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame famOther 
      Caption         =   "Also Place in Letter"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox chkMyAddress 
         Caption         =   "My Address"
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
         TabIndex        =   18
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkAddressLetter 
         Caption         =   "Addresses Address"
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
         TabIndex        =   17
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "Todays Date"
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
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Signing"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton optSigning 
         Caption         =   "Love"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton optSigning 
         Caption         =   "Yours Truly"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSigning 
         Caption         =   "Sincerely"
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
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
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
      Left            =   3480
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
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
      Left            =   4800
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtMyAddress 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtAddresse 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Frame famLetterType 
      Caption         =   "Type of Letter"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optBusiness 
         Caption         =   "Business Letter"
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
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optPersonal 
         Caption         =   "Personal Letter"
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
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Left            =   6000
      TabIndex        =   22
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Letter to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   21
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type any changes to your address here:"
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
      TabIndex        =   12
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label lblAddresse 
      BackStyle       =   0  'Transparent
      Caption         =   "Addresse:"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "LetterWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAddressLetter_Click()
    ShowQuickPreview
End Sub

Private Sub chkDate_Click()
    ShowQuickPreview
End Sub

Private Sub chkMyAddress_Click()
    famPlaceMyAddress.Visible = chkMyAddress.value
    ShowQuickPreview
End Sub

Private Sub cmdCancel_Click()
 SelFormat.SFormat = ""
 Unload Me
End Sub

Private Sub cmdHelp_Click()
    NoHelp
End Sub

Private Sub cmdOK_Click()
 SaveSettings
 SelFormat.SFormat = CreateLetter()
 Unload Me
End Sub

Private Sub cmdPreview_Click()
 MsgBox CreateLetter(), , "Preview of Letter"
End Sub

Private Function CreateLetter() As String
  ' Creates a letter based on the selected items
  Dim Letter As String, i As Integer
  SetCRLF

  If chkMyAddress And optPlaceMyAddress(0) Then Letter = txtMyAddress & crlf & crlf
  
  If chkDate Then Letter = Letter & Format$(Date$, "dddd, mmmm, d, yyyy") & crlf & crlf
  
  If chkAddressLetter Then Letter = Letter & txtAddresse & crlf & crlf

  Letter = Letter & "Dear "
  If Trim$(txtSendTo) <> "" Then
      Letter = Letter & Trim$(txtSendTo)
  Else
      Letter = Letter & "[Addresses Name Here]"
  End If

  If optPersonal Then
      Letter = Letter & ","
  Else
      Letter = Letter & ":"
  End If
  Letter = Letter & crlf & crlf & "[Type the body of the letter here]" & crlf & crlf

  For i = 0 To 2
      If optSigning(i) Then Letter = Letter & optSigning(i).Caption & ","
  Next i
  
  If chkMyAddress And optPlaceMyAddress(1) Then Letter = Letter & crlf & crlf & crlf & txtMyAddress
  
  CreateLetter = Letter
End Function

Private Sub famLetterType_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
'frmMDI.lblStatus = "Specifies the type of the letter"
End Sub

Private Sub Form_Load()
 LoadSettings
 CenterMe Me
End Sub


Private Sub LoadSettings()
 ' Load the last saved settings
 Dim Fetch As String
 
 ' Load My Address
 Fetch = GetMyIni("BDMAN.INI", "Letter Wizard", "MyAddress")
 Fetch = BracesToCRLF(Fetch)
 If Trim$(Fetch) = "" Then
    txtMyAddress = GetWinIni("BDMAN", "UserName")
 Else
    txtMyAddress = Fetch
 End If
 
 ' Load Addresse Address
 Fetch = GetMyIni("BDMAN.INI", "Letter Wizard", "Addresse")
 txtAddresse = BracesToCRLF(Fetch)

 ' Load Send To Info
 txtSendTo = GetMyIni("BDMAN.INI", "Letter Wizard", "SendTo")

 
 If Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "SettingsSaved")) <> "" Then
 ' Load the letter type info
 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "PBLetter"))
 optPersonal.value = Not Val(Fetch)
 optBusiness.value = Val(Fetch)
 If Fetch = "0" Then optSigning(2).Visible = False

 ' Load the signing info
 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "Sincerely"))
 optSigning(0).value = Val(Fetch)
 
 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "Yours Truly"))
 optSigning(1).value = Val(Fetch)

 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "Love"))
 optSigning(2).value = Val(Fetch)

 ' Load the other settings
 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "Todays Date"))
 chkDate.value = Val(Fetch)

 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "AddresseAddress"))
 chkAddressLetter.value = Val(Fetch)

 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "UseMyAddress"))
 chkMyAddress.value = Val(Fetch)
 famPlaceMyAddress.Visible = Val(Fetch)

 Fetch = Trim$(GetMyIni("BDMAN.INI", "Letter Wizard", "PlaceMyAddress"))
 optPlaceMyAddress(0).value = Not Val(Fetch)
 optPlaceMyAddress(1).value = Val(Fetch)
 End If
 
End Sub

Private Sub optBusiness_Click()
    optSigning(2).Visible = False
    If optSigning(2).value Then optSigning(0) = 1
    ShowQuickPreview
End Sub

Private Sub optPersonal_Click()
    optSigning(2).Visible = True
    ShowQuickPreview
End Sub

Private Sub optPlaceMyAddress_Click(Index As Integer)
    ShowQuickPreview
End Sub

Private Sub optSigning_Click(Index As Integer)
    ShowQuickPreview
End Sub

Private Sub SaveSettings()
 ' Save  the options the user selected.
 Dim Setting As String
 SetMyIni "BDMAN.INI", "Letter Wizard", "SettingsSaved", "Yes"
 
 ' Save my address
 Setting = txtMyAddress
 Setting = CRLFToBraces(Setting)
 SetMyIni "BDMAN.INI", "Letter Wizard", "MyAddress", Setting

 ' Save addresse address
 Setting = txtAddresse
 Setting = CRLFToBraces(Setting)
 SetMyIni "BDMAN.INI", "Letter Wizard", "Addresse", Setting

 ' Save the sent to info
 Setting = txtSendTo
 Setting = CRLFToBraces(Setting)
 SetMyIni "BDMAN.INI", "Letter Wizard", "SendTo", Setting

 ' Save the letter type.  If the optBusiness is false then the letter is
 ' a personal letter.
 Setting = Trim$(Str(optBusiness.value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "PBLetter", Setting

 ' Save signing settings
 Setting = Trim$(Str(optSigning(0).value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "Sincerely", Setting
 
 Setting = Trim$(Str(optSigning(1).value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "Yours Truly", Setting
 
 Setting = Trim$(Str(optSigning(2).value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "Love", Setting

 ' Save the other settings
 Setting = Trim$(Str(chkDate.value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "Todays Date", Setting

 Setting = Trim$(Str(chkAddressLetter.value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "AddresseAddress", Setting

 Setting = Trim$(Str(chkMyAddress.value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "UseMyAddress", Setting

 Setting = Trim$(Str(optPlaceMyAddress(1).value))
 SetMyIni "BDMAN.INI", "Letter Wizard", "PlaceMyAddress", Setting

End Sub

Private Sub ShowQuickPreview()
    Dim p As String, t As String, i As Integer

    p = CreateLetter()
    lblPreview = " "
    For i = 1 To Len(p)
        t = Mid$(p, i, 1)
        If t = Chr$(13) Then
            lblPreview = lblPreview & crlf & " "
        ElseIf t = Chr$(10) Then

        Else
            lblPreview = lblPreview & t
        End If
    Next i
End Sub

Private Sub txtAddresse_Change()
    ShowQuickPreview
End Sub

Private Sub txtMyAddress_Change()
    ShowQuickPreview
End Sub

Private Sub txtSendTo_Change()
    ShowQuickPreview
End Sub

