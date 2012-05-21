VERSION 5.00
Begin VB.Form frmPrefrences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prefrences"
   ClientHeight    =   2385
   ClientLeft      =   1455
   ClientTop       =   2445
   ClientWidth     =   5640
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
   ScaleHeight     =   2385
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame famModifyDialog 
      Caption         =   "Modify Dialog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdSetFieldDefaults 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Set"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkFillData 
         Caption         =   "Set Field Data to Default Data in new Items"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CheckBox chkShowISPButton 
         Caption         =   "Show Internet Service Provider Button"
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
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkShowStateButton 
         Caption         =   "Show Select State Abbreviation Button"
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
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   3975
      End
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
      Left            =   4200
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrefrences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TempData As FolderType

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim PString As String
    Prefrences.ShowStateButton = chkShowStateButton.Value
    PString = Trim$(Str$(Prefrences.ShowStateButton))
    Prefrences.ShowISPButton = chkShowISPButton.Value
    PString = PString & Trim$(Str$(Prefrences.ShowISPButton))
    
    If FolderOpen Then
        Prefrences.FillData = chkFillData.Value
        PString = PString & Trim$(Str$(Prefrences.FillData))

        DataToFill.Company = TempData.Company
        DataToFill.AName = TempData.AName
        DataToFill.Address = TempData.Address
        DataToFill.City = TempData.City
        DataToFill.State = TempData.State
        DataToFill.Zip_Code = TempData.Zip_Code
        DataToFill.Home_Phone = TempData.Home_Phone
        DataToFill.Bus_Phone = TempData.Bus_Phone
        DataToFill.Pager = TempData.Pager
        DataToFill.Fax = TempData.Fax
        DataToFill.E_Mail = TempData.E_Mail
        DataToFill.WebPage = TempData.WebPage
        DataToFill.Notes = TempData.Notes
        
        Pd.Company = TempData.Company
        Pd.AName = TempData.AName
        Pd.Address = TempData.Address
        Pd.City = TempData.City
        Pd.State = TempData.State
        Pd.Zip_Code = TempData.Zip_Code
        Pd.Home_Phone = TempData.Home_Phone
        Pd.Bus_Phone = TempData.Bus_Phone
        Pd.Pager = TempData.Pager
        Pd.Fax = TempData.Fax
        Pd.E_Mail = TempData.E_Mail
        Pd.WebPage = TempData.WebPage
        Pd.Notes = TempData.Notes
        PutPD FreeNum, 114, True
    End If
    GetPD FreeNum, 1, False
    Pd.Home_Phone = PString
    PutPD FreeNum, 1, False
    Unload Me
End Sub

Private Sub cmdSetFieldDefaults_Click()
 ResetPD
 Pd.Company = TempData.Company
 Pd.AName = TempData.AName
 Pd.Address = TempData.Address
 Pd.City = TempData.City
 Pd.State = TempData.State
 Pd.Zip_Code = TempData.Zip_Code
 Pd.Home_Phone = TempData.Home_Phone
 Pd.Bus_Phone = TempData.Bus_Phone
 Pd.Pager = TempData.Pager
 Pd.Fax = TempData.Fax
 Pd.E_Mail = TempData.E_Mail
 Pd.WebPage = TempData.WebPage
 Pd.Notes = TempData.Notes


 Load EditWindow
 EditWindow.Caption = "Prefrences:  Default Field Data"
 EditWindow.cmdBookmark.Visible = False
 EditWindow.Command2.Left = 5880
 EditWindow.Command1.Left = 4560
 EditWindow.Command1.Caption = "&Save"
 EditWindow.Show 1
 If Not Canceled Then
    TempData.Company = Pd.Company
    TempData.AName = Pd.AName
    TempData.Address = Pd.Address
    TempData.City = Pd.City
    TempData.State = Pd.State
    TempData.Zip_Code = Pd.Zip_Code
    TempData.Home_Phone = Pd.Home_Phone
    TempData.Bus_Phone = Pd.Bus_Phone
    TempData.Pager = Pd.Pager
    TempData.Fax = Pd.Fax
    TempData.E_Mail = Pd.E_Mail
    TempData.WebPage = Pd.WebPage
    TempData.Notes = Pd.Notes
    
    chkFillData.Value = 1
 End If
End Sub

Private Sub Form_Load()
    TempData.Company = DataToFill.Company
    TempData.AName = DataToFill.AName
    TempData.Address = DataToFill.Address
    TempData.City = DataToFill.City
    TempData.State = DataToFill.State
    TempData.Zip_Code = DataToFill.Zip_Code
    TempData.Home_Phone = DataToFill.Home_Phone
    TempData.Bus_Phone = DataToFill.Bus_Phone
    TempData.Pager = DataToFill.Pager
    TempData.Fax = DataToFill.Fax
    TempData.E_Mail = DataToFill.E_Mail
    TempData.WebPage = DataToFill.WebPage
    TempData.Notes = DataToFill.Notes

    chkShowStateButton.Value = Prefrences.ShowStateButton
    chkShowISPButton.Value = Prefrences.ShowISPButton
    chkFillData.Value = Prefrences.FillData
End Sub

