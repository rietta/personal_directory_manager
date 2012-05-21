VERSION 5.00
Begin VB.Form ChangeActiveTemplateDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Active Template"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "ChangeActiveTemplateDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ListBox lstTemplates 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Templates:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"ChangeActiveTemplateDialog.frx":0442
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "ChangeActiveTemplateDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdOK_Click()
    CurrentFolderTemplate = lstTemplates.ListIndex
    SetTemplateFieldsToActive CurrentFolderTemplate
    SaveFields
    Canceled = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    lstTemplates.Clear
    For i = 0 To UBound(TemplateIndex)
        lstTemplates.AddItem TemplateIndex(i)
    Next i
    Canceled = True
End Sub
