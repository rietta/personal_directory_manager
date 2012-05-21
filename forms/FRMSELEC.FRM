VERSION 5.00
Begin VB.Form frmSelectItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Items"
   ClientHeight    =   4635
   ClientLeft      =   1170
   ClientTop       =   1545
   ClientWidth     =   4170
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
   ScaleHeight     =   4635
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
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
      Left            =   2880
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company:                             Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmSelectItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 Dim i As Integer, Num As Integer
 Num = 0
 ReDim SelectedItems(0) As Integer
 For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
         ReDim Preserve SelectedItems(Num) As Integer
         SelectedItems(Num) = i + 2
         Num = Num + 1
    End If
 Next i
 Unload Me
End Sub

Private Sub Form_Load()
     On Error Resume Next
     Dim j As Integer
     List1.Clear
     DoTitle Me.Label1
      
     For j = 2 To 100
        Get #FreeNum, j, Pd
        Pd.Company = Decript(Trim$(Pd.Company))
        Pd.AName = Decript(Trim$(Pd.AName))
        If RTrim$(Pd.AName) <> "" Or RTrim$(Pd.Company) <> "" Then
          List1.AddItem Trim$(Pd.Company) & Chr$(9) & Chr$(9) & Trim$(Pd.AName)
        End If
     Next j
End Sub

