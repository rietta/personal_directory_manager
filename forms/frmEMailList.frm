VERSION 5.00
Begin VB.Form frmEMailList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail Mailing List Generator"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmEMailList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy List to Clipboard"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate List"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtMailingList 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CheckBox chkUseAOL 
      Caption         =   "I am going to send this using AOL's e-mail system."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEMailList.frx":030A
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmEMailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.SetText txtMailingList
End Sub

' Written on 6/22/99
Private Sub cmdGenerate_Click()
    Dim i As Integer, usingAOL As Boolean
    Dim strList As String, strEMail As String
    
    If chkUseAOL.value = vbChecked Then
        usingAOL = True
    Else
        usingAOL = False
    End If
    
    For i = 2 To 100
        Get #FreeNum, i, Pd
        strEMail = Decript(Trim$(Pd.E_Mail))
        If strEMail <> "" Then
            If strList <> "" Then strList = strList + vbCrLf
            
            If usingAOL Then
                strList = strList + "(" + strEMail + ")"
            Else
                strList = strList + strEMail
            End If
        End If
    Next i
    
    If strList <> "" Then txtMailingList = strList
End Sub

