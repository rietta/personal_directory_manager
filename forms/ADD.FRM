VERSION 5.00
Begin VB.Form EditWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add"
   ClientHeight    =   3975
   ClientLeft      =   690
   ClientTop       =   2055
   ClientWidth     =   8250
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
   ScaleHeight     =   3975
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Clear"
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
      TabIndex        =   31
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectISP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "|"
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
      Left            =   8040
      TabIndex        =   14
      Top             =   2280
      Width           =   135
   End
   Begin VB.CommandButton cmdBookmark 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Add &Bookmark"
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
      Left            =   6360
      TabIndex        =   30
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelState 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "|"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   1200
      Width           =   135
   End
   Begin VB.TextBox txtNotes 
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
      Left            =   1680
      MaxLength       =   100
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   3000
      Width           =   6495
   End
   Begin VB.TextBox txt_WebPage 
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
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   6495
   End
   Begin VB.TextBox txt_Company 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txt_BusAddress 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txt_Email 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2280
      Width           =   6375
   End
   Begin VB.TextBox txt_Fax 
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
      Left            =   5520
      MaxLength       =   13
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txt_Pager 
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
      Left            =   1680
      MaxLength       =   13
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txt_Name 
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
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txt_Busphone 
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
      Left            =   5520
      MaxLength       =   20
      TabIndex        =   8
      Text            =   " "
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txt_HomePhone 
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
      Left            =   1680
      MaxLength       =   13
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txt_BusZipCode 
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
      Left            =   6840
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txt_BusState 
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
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txt_BusCity 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   1935
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
      Left            =   5040
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Save"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
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
      Index           =   12
      Left            =   0
      TabIndex        =   17
      Tag             =   "NO 3D"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page (URL):"
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
      Index           =   11
      Left            =   0
      TabIndex        =   19
      Tag             =   "NO 3D"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
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
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Tag             =   "NO 3D"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
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
      Index           =   9
      Left            =   4200
      TabIndex        =   21
      Tag             =   "NO 3D"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pager:"
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
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Tag             =   "NO 3D"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bus. Phone:"
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
      Index           =   7
      Left            =   4200
      TabIndex        =   23
      Tag             =   "NO 3D"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Tag             =   "NO 3D"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
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
      Index           =   5
      Left            =   6000
      TabIndex        =   25
      Tag             =   "NO 3D"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Index           =   4
      Left            =   3720
      TabIndex        =   26
      Tag             =   "NO 3D"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Tag             =   "NO 3D"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Tag             =   "NO 3D"
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Tag             =   "NO 3D"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Tag             =   "NO 3D"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "EditWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBookmark_Click()
 If cmdBookmark.Caption = "Add Bookmark" Then
    cmdBookmark.Caption = "Remove Bookmark"
 Else
    cmdBookmark.Caption = "Add Bookmark"
 End If
 Command1.Enabled = True
End Sub

Private Sub cmdClear_Click()
   txt_Company = ""
   txt_Name = ""
   txt_BusAddress = ""
   txt_BusCity = ""
   txt_BusState = ""
   txt_BusZipCode = ""
   txt_HomePhone = ""
   txt_Busphone = ""
   txt_Pager = ""
   txt_Fax = ""
   txt_Email = ""
   txt_WebPage.text = ""
   txtNotes.text = ""

   Command1.Enabled = False
End Sub

Private Sub cmdHelp_Click()
Select Case UCase(Caption)
    Case "ADD ITEM": 'ShowHelpTopic 10
    Case "MODIFY ITEM": 'ShowHelpTopic 11
End Select
End Sub

Private Sub cmdSelectISP_Click()
 Dim ISP As String
 ISP = GetIP()      ' Allow the user to select an ISP
 If ISP <> "" Then txt_Email.SelText = ISP
End Sub

Private Sub cmdSelState_Click()
 Screen.MousePointer = 11
 On Error Resume Next
 frmStates.Show 1
 If SelectedState <> "" Then txt_BusState.text = Trim$(SelectedState)
End Sub

Private Sub Command1_Click()
   ResetPD
   Canceled = False
   Pd.Company = Trim$(txt_Company)
   Pd.AName = Trim$(txt_Name)
   Pd.Address = Trim$(txt_BusAddress)
   Pd.City = Trim$(txt_BusCity)
   Pd.State = Trim$(txt_BusState)
   Pd.Zip_Code = Trim$(txt_BusZipCode)
   Pd.Home_Phone = Trim$(txt_HomePhone)
   Pd.Bus_Phone = Trim$(txt_Busphone)
   Pd.Pager = Trim$(txt_Pager)
   Pd.Fax = Trim$(txt_Fax)
   Pd.E_Mail = Trim$(txt_Email)
   Pd.WebPage = Trim$(txt_WebPage.text)
   Pd.Notes = Trim$(txtNotes.text)
   
   If cmdBookmark.Caption = "Remove Bookmark" Then
      Pd.Bookmark = True
   Else
      Pd.Bookmark = False
   End If
   
   Unload Me
End Sub

Private Sub Command2_Click()
 ResetPD
 Canceled = True
 Unload Me
End Sub

Private Sub Form_Load()
   CenterMe Me

   cmdSelState.Visible = Prefrences.ShowStateButton
   If Prefrences.ShowISPButton = 0 Then
        cmdSelectISP.Visible = False
        txt_Email.Width = 6495
   End If

   '-----------------------------------------------------
   'Set max lengths for differnt Fields
   '-----------------------------------------------------
   
   txt_Company.MaxLength = Len(Pd.Company)
   txt_Name.MaxLength = Len(Pd.AName)
   txt_BusAddress.MaxLength = Len(Pd.Address)
   txt_BusCity.MaxLength = Len(Pd.City)
   txt_BusState.MaxLength = Len(Pd.State)
   txt_BusZipCode.MaxLength = Len(Pd.Zip_Code)
   txt_HomePhone.MaxLength = Len(Pd.Home_Phone)
   txt_Busphone.MaxLength = Len(Pd.Bus_Phone)
   txt_Pager.MaxLength = Len(Pd.Pager)
   txt_Fax.MaxLength = Len(Pd.Fax)
   txt_Email.MaxLength = Len(Pd.E_Mail)
   txt_WebPage.MaxLength = Len(Pd.WebPage)
   txtNotes.MaxLength = Len(Pd.Notes)
   
   '-----------------------------------------------------
   'Set current data into fields
   '-----------------------------------------------------
   
   txt_Company = Trim$(Pd.Company)
   txt_Name = Trim$(Pd.AName)
   txt_BusAddress = Trim$(Pd.Address)
   txt_BusCity = Trim$(Pd.City)
   txt_BusState = Trim$(Pd.State)
   txt_BusZipCode = Trim$(Pd.Zip_Code)
   txt_HomePhone = Trim$(Pd.Home_Phone)
   txt_Busphone = Trim$(Pd.Bus_Phone)
   txt_Pager = Trim$(Pd.Pager)
   txt_Fax = Trim$(Pd.Fax)
   txt_Email = Trim$(Pd.E_Mail)
   txt_WebPage = Trim$(Pd.WebPage)
   txtNotes = Trim$(Pd.Notes)
 
   If Pd.Bookmark Then cmdBookmark.Caption = "Remove &Bookmark"
   '----------------------------------------------------
   ' Load field names for user definable fields  if the field is blank
   ' then hide the field from view
   '----------------------------------------------------
   
   For i% = 0 To 12
      If UserField(i%) <> "" Then
           lblField(i%) = UserField(i%) & ":"
      Else
           lblField(i%).Visible = False
      End If
   Next i%
   
   If UserField(0) = "" Then
       txt_Company.Visible = False
       txt_Company.Tag = "NO 3D"
   End If
   If UserField(1) = "" Then
       txt_Name.Visible = False
       txt_Name.Tag = "NO 3D"
   End If
   If UserField(2) = "" Then
       txt_BusAddress.Visible = False
       txt_BusAddress.Tag = "NO 3D"
   End If
   If UserField(3) = "" Then
       txt_BusCity.Visible = False
       txt_BusCity.Tag = "NO 3D"
   End If
   If UserField(4) = "" Then
       txt_BusState.Visible = False
       txt_BusState.Tag = "NO 3D"
       cmdSelState.Visible = False
       cmdSelState.Tag = "NO 3D"
   End If
   If UserField(5) = "" Then
       txt_BusZipCode.Visible = False
       txt_BusZipCode.Tag = "NO 3D"
   End If
   If UserField(6) = "" Then
       txt_HomePhone.Visible = False
       txt_HomePhone.Tag = "NO 3D"
   End If
   If UserField(7) = "" Then
       txt_Busphone.Visible = False
       txt_Busphone.Tag = "NO 3D"
   End If
   If UserField(8) = "" Then
       txt_Pager.Visible = False
       txt_Pager.Tag = "NO 3D"
   End If
   If UserField(9) = "" Then
       txt_Fax.Visible = False
       txt_Fax.Tag = "NO 3D"
   End If
   If UserField(10) = "" Then
       txt_Email.Visible = False
       txt_Email.Tag = "NO 3D"
       cmdSelectISP.Visible = False
   End If
   If UserField(11) = "" Then
       txt_WebPage.Visible = False
       txt_WebPage.Tag = "NO 3D"
   End If
   If UserField(12) = "" Then
       txtNotes.Visible = False
       txtNotes.Tag = "NO 3D"
   End If
   '----------------------------------------------------
   'Final details
   '----------------------------------------------------
   
   EditWindow.Caption = Si$    'Set form caption 'Add or Modify
   Command1.Enabled = False        'Disable continue button

   Screen.MousePointer = 0

End Sub

Private Sub mnuStateArray_Click(Index As Integer)
 Dim s As String
 Select Case Index
     Case 0: s = "AL"
     Case 1: s = "AK"
     Case 2: s = "AZ"
     Case 3: s = "AR"
     Case 4: s = "CA"
     Case 5: s = "CO"
     Case 6: s = "CT"
     Case 7: s = "DE"
     Case 8: s = "FL"
     Case 9: s = "GA"
     Case 10: s = "HI"
     Case 11: s = "ID"
     Case 12: s = "IL"
     Case 13: s = "IN"
     Case 14: s = "IA"
     Case 15: s = "KS"
     Case 16: s = "KY"
     Case 17: s = "LA"
     Case 18: s = "ME"
     Case 19: s = "MA"
     Case 20: s = "MI"
     Case 21: s = "MN"
     Case 22: s = "MS"
     Case 23: s = "MO"
     Case 24: s = "MT"
     Case 25: s = "NE"
     Case 26: s = "NV"
     Case 27: s = "NH"
     Case 28: s = "NJ"
     Case 29: s = "NM"
     Case 30: s = "NY"
     Case 31: s = "NC"
     Case 32: s = "ND"
     Case 33: s = "OH"
     Case 34: s = "OK"
     Case 35: s = "OR"
     Case 36: s = "PA"
     Case 37: s = "RI"
     Case 38: s = "SC"
     Case 39: s = "SD"
     Case 40: s = "TN"
     Case 41: s = "TX"
     Case 42: s = "UT"
     Case 43: s = "VT"
     Case 44: s = "VA"
     Case 45: s = "WA"
     Case 46: s = "WV"
     Case 47: s = "WI"
     Case 48: s = "WY"
     Case 49: s = "ZZ"
End Select
txt_BusState.text = s
End Sub

Private Sub text_ZipCode_Change()
Pd.Zip_Code = txt_ZipCode
End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text6_Change()

End Sub

Private Sub Text7_Change()

End Sub

Private Sub txt_BusAddress_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_BusAddress_GotFocus()
    txt_BusAddress.SelStart = 0
    txt_BusAddress.SelLength = txt_BusAddress.MaxLength
End Sub


Private Sub txt_BusAddress_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_BusCity_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True
End Sub

Private Sub txt_BusCity_GotFocus()
    txt_BusCity.SelStart = 0
    txt_BusCity.SelLength = txt_BusCity.MaxLength
End Sub


Private Sub txt_BusCity_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_Busphone_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_Busphone_GotFocus()
    txt_Busphone.SelStart = 0
    txt_Busphone.SelLength = txt_Busphone.MaxLength
End Sub


Private Sub txt_BusState_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True
End Sub

Private Sub txt_BusState_GotFocus()
    txt_BusState.SelStart = 0
    txt_BusState.SelLength = txt_BusState.MaxLength
End Sub


Private Sub txt_BusState_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_BusZipCode_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_BusZipCode_GotFocus()
    txt_BusZipCode.SelStart = 0
    txt_BusZipCode.SelLength = txt_BusZipCode.MaxLength
End Sub


Private Sub txt_BusZipCode_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_Company_Change()
    TC$ = txt_Company & txt_Name
    If Caption <> "Prefrences:  Default Field Data" Then
        If Trim$(TC$) <> "" Then
            Command1.Enabled = True
        Else
            Command1.Enabled = False
        End If
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub txt_Company_GotFocus()
    txt_Company.SelStart = 0
    txt_Company.SelLength = txt_Company.MaxLength
End Sub


Private Sub txt_Company_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_Email_Change()
    Dim of As Integer, resp As Integer
   
    TC$ = txt_Company & txt_Name
    If Trim$(TC$) <> "" Then Command1.Enabled = True
   
    of = InStr(LCase(txt_Email), "mailto:")
    If of >= 1 Then
        resp = MsgBox("The HTML command 'mailto' should not be placed in the e-mail box.  Do you want me to remove it for you?", vbYesNo + vbInformation, SPROGRAMNAME)
        If resp = vbYes Then
            txt_Email.SelStart = of - 1
            txt_Email.SelLength = Len("mailto:")
            If LCase(txt_Email.SelText) = "mailto:" Then txt_Email.SelText = ""
        End If
    End If

End Sub

Private Sub txt_Email_GotFocus()
    txt_Email.SelStart = 0
    txt_Email.SelLength = txt_Email.MaxLength
End Sub


Private Sub txt_Email_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_Fax_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_Fax_GotFocus()
    txt_Fax.SelStart = 0
    txt_Fax.SelLength = txt_Fax.MaxLength
End Sub


Private Sub txt_Fax_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_HomePhone_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_HomePhone_GotFocus()
    txt_HomePhone.SelStart = 0
    txt_HomePhone.SelLength = txt_HomePhone.MaxLength
End Sub


Private Sub txt_HomePhone_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_LastName_Change()
If LTrim$(txt_LastName) <> "" Then Command1.Enabled = True
End Sub

Private Sub txt_Name_Change()
   TC$ = txt_Company & txt_Name
   If Caption <> "Prefrences:  Default Field Data" Then
     If Trim$(TC$) <> "" Then
           Command1.Enabled = True
     Else
         Command1.Enabled = False
     End If
   Else
        Command1.Enabled = True
   End If
End Sub

Private Sub txt_Name_GotFocus()
    txt_Name.SelStart = 0
    txt_Name.SelLength = txt_Name.MaxLength
End Sub


Private Sub txt_Name_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_Pager_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_Pager_GotFocus()
    txt_Pager.SelStart = 0
    txt_Pager.SelLength = txt_Pager.MaxLength
End Sub


Private Sub txt_Pager_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

Private Sub txt_WebPage_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txt_WebPage_GotFocus()
    txt_WebPage.SelStart = 0
    txt_WebPage.SelLength = txt_WebPage.MaxLength
End Sub


Private Sub txt_WebPage_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0
End Sub

Private Sub txtNotes_Change()
   TC$ = txt_Company & txt_Name
   If Trim$(TC$) <> "" Then Command1.Enabled = True

End Sub

Private Sub txtNotes_GotFocus()
    txtNotes.SelStart = 0
    txtNotes.SelLength = txtNotes.MaxLength
End Sub


Private Sub txtNotes_KeyPress(KeyAscii As Integer)
 If Chr$(KeyAscii) = "&" Then KeyAscii = 0  'Check for valid charactors
End Sub

