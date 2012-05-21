VERSION 5.00
Begin VB.Form FirstTimeUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Time User"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FirstTimeUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtOrganization 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FirstTimeUser.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Organization:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"FirstTimeUser.frx":0884
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FirstTimeUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    UserInformation.UName = txtName
    UserInformation.UCompany = txtOrganization
    SetMyIni "ERSOFT.INI", "User", "Name", UserInformation.UName
    SetMyIni "ERSOFT.INI", "User", "Company", UserInformation.UCompany
    Unload Me
End Sub

Private Sub Form_Load()
    txtName = UserInformation.UName
    txtOrganization = UserInformation.UCompany
    txtName_Change
End Sub

Private Sub txtName_Change()
    If Trim$(txtName) <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
