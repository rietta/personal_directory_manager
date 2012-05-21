VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PrintStyles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Binder Properties"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "PrintStyles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Columns"
      Height          =   855
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
      Begin VB.OptionButton optColumn 
         Caption         =   "1 Column"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optColumn 
         Caption         =   "2 Columns"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Timer tmrFirstShow 
      Interval        =   1000
      Left            =   3000
      Top             =   4560
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3960
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Apply Changes"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CheckBox chkPageNums 
      Caption         =   "Print Page Numbers"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdBody 
      Caption         =   "Change Body Style"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdHeader 
      Caption         =   "Change Header Style"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Line lneDivider 
         Visible         =   0   'False
         X1              =   2160
         X2              =   2160
         Y1              =   3840
         Y2              =   480
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   0
         Picture         =   "PrintStyles.frx":0442
         Top             =   3840
         Width           =   4395
      End
      Begin VB.Label lblBody 
         BackStyle       =   0  'Transparent
         Caption         =   "Body Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Binder Title - Page 1                                 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"PrintStyles.frx":5A74
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   4455
   End
End
Attribute VB_Name = "PrintStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nWRatio As Double, nHRatio As Double, nOverallRatio As Double

Dim bFontDialogOpen As Boolean

Private Sub chkPageNums_Click()
    GenerateScreen
End Sub

Private Sub cmdBody_Click()
    
    If bFontDialogOpen Then Exit Sub
    bFontDialogOpen = True
    
   cmdlg.FontName = lblBody.Font.Name
   cmdlg.FontSize = lblBody.Font.size
   cmdlg.FontBold = lblBody.Font.Bold
   cmdlg.FontItalic = lblBody.Font.Italic
   cmdlg.FontUnderline = lblBody.Font.Underline
   cmdlg.FontStrikethru = lblBody.FontStrikethru
   cmdlg.color = lblBody.ForeColor
   
   
   ' Set Cancel to True.
   cmdlg.CancelError = True
     
   On Error GoTo ErrHandler
   ' Set the Flags property.
   cmdlg.FLAGS = cdlCFBoth Or cdlCFEffects
   ' Display the Font dialog box.
   cmdlg.ShowFont
   
   ' Set text properties according to user's
   ' selections.
   lblBody.Font.Name = cmdlg.FontName
   lblBody.Font.size = cmdlg.FontSize
   lblBody.Font.Bold = cmdlg.FontBold
   lblBody.Font.Italic = cmdlg.FontItalic
   lblBody.Font.Underline = cmdlg.FontUnderline
   lblBody.FontStrikethru = cmdlg.FontStrikethru
   lblBody.ForeColor = cmdlg.color
 
ErrHandler:
   ' User pressed Cancel button.
   
   bFontDialogOpen = False
   
   Exit Sub

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHeader_Click()
    If bFontDialogOpen Then Exit Sub
    bFontDialogOpen = True
    
   cmdlg.FontName = lblHeader.Font.Name
   cmdlg.FontSize = lblHeader.Font.size
   cmdlg.FontBold = lblHeader.Font.Bold
   cmdlg.FontItalic = lblHeader.Font.Italic
   cmdlg.FontUnderline = lblHeader.Font.Underline
   cmdlg.FontStrikethru = lblHeader.FontStrikethru
   cmdlg.color = lblHeader.ForeColor
   
   
   ' Set Cancel to True.
   cmdlg.CancelError = True
     
   On Error GoTo ErrHandler
   ' Set the Flags property.
   cmdlg.FLAGS = cdlCFBoth Or cdlCFEffects
   ' Display the Font dialog box.
   cmdlg.ShowFont
   
   ' Set text properties according to user's
   ' selections.
   lblHeader.Font.Name = cmdlg.FontName
   lblHeader.Font.size = cmdlg.FontSize
   lblHeader.Font.Bold = cmdlg.FontBold
   lblHeader.Font.Italic = cmdlg.FontItalic
   lblHeader.Font.Underline = cmdlg.FontUnderline
   lblHeader.FontStrikethru = cmdlg.FontStrikethru
   lblHeader.ForeColor = cmdlg.color
   

ErrHandler:
   ' User pressed Cancel button.
   bFontDialogOpen = False
   Exit Sub

End Sub

Private Sub cmdOK_Click()
    HeadingProp.sName = lblHeader.Font.Name
    HeadingProp.nSize = lblHeader.Font.size
    HeadingProp.bBold = lblHeader.Font.Bold
    HeadingProp.bItalic = lblHeader.Font.Italic
    HeadingProp.bUnderline = lblHeader.Font.Underline
    HeadingProp.bStrikethru = lblHeader.FontStrikethru
    HeadingProp.nColor = lblHeader.ForeColor
    
    HeadingProp.bPageNumbers = chkPageNums
    
    BodyProp.sName = lblBody.Font.Name
    BodyProp.nSize = lblBody.Font.size
    BodyProp.bBold = lblBody.Font.Bold
    BodyProp.bItalic = lblBody.Font.Italic
    BodyProp.bUnderline = lblBody.Font.Underline
    BodyProp.bStrikethru = lblBody.FontStrikethru
    BodyProp.nColor = lblBody.ForeColor
    
    If optColumn(0) Then
        BodyProp.nColumns = 1
    Else
        BodyProp.nColumns = 2
    End If
    
    SavePrintFormatting
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    bFontDialogOpen = False
        
    lblHeader.Font.Name = HeadingProp.sName
    lblHeader.Font.size = HeadingProp.nSize
    lblHeader.Font.Bold = HeadingProp.bBold
    lblHeader.Font.Italic = HeadingProp.bItalic
    lblHeader.Font.Underline = HeadingProp.bUnderline
    lblHeader.FontStrikethru = HeadingProp.bStrikethru
    lblHeader.ForeColor = HeadingProp.nColor
    
    If HeadingProp.bPageNumbers Then chkPageNums.value = vbChecked
    
    lblBody.Font.Name = BodyProp.sName
    lblBody.Font.size = BodyProp.nSize '* nOverallRatio
    lblBody.Font.Bold = BodyProp.bBold
    lblBody.Font.Italic = BodyProp.bItalic
    lblBody.Font.Underline = BodyProp.bUnderline
    lblBody.FontStrikethru = BodyProp.bStrikethru
    lblBody.ForeColor = BodyProp.nColor
    
    If BodyProp.nColumns = 2 Then
        optColumn(1) = True
    Else
        optColumn(0) = True
    End If
    
    DoEvents
    GenerateScreen
    
End Sub

Private Sub lblBody_Click()
    cmdBody_Click
End Sub

Private Sub lblHeader_Click()
    cmdHeader_Click
End Sub

Public Sub GenerateScreen()
    
    lblHeader = OpenFolderName
    
    If chkPageNums Then lblHeader = lblHeader + " - Page 1"
    
    ' Generate Body
    
    If optColumn(1) Then
        lneDivider.Visible = True
        lblBody.Width = lneDivider.X1 - lblBody.Left * 2
    Else
        lneDivider.Visible = False
        lblBody.Width = picPreview.ScaleWidth - lblBody.Width
    End If
    
    lblBody = ""
    GetPD FreeNum, SelectedItems(0), True
    PrintCurrentData
End Sub

Private Sub PrintCurrentData()
'This procedure prints the simulates the printing
' done by the actual print function

'-------------------------------------------------

Dim NumItems As Integer, i   As Integer

ReDim PrintThisField(10) As Integer

For i = 0 To 10
    If PrintBinderDialog.chkAddvancedOptions(i) Then
        PrintThisField(i) = True
    Else
        PrintThisField(i) = False
    End If
Next i

 '------------------------------------------------------------------------
 'Determin which fields wont be printed
 '------------------------------------------------------------------------
 
 If Not PrintBinderDialog.chkBlank Then
      'Find which fields are blank so they won't be printed.
      If Trim(Pd.Company) = "" Then
          NumItems = NumItems - 1
          PrintThisField(0) = False
      End If
      If Trim(Pd.AName) = "" Then
          NumItems = NumItems - 1
          PrintThisField(1) = False
      End If
      If Trim(Pd.Address) = "" Then
          NumItems = NumItems - 1
          PrintThisField(2) = False
      End If
      If (Trim(Pd.City) = "") And (Trim(Pd.State) = "") And (Trim(Pd.Zip_Code) = "") Then
          NumItems = NumItems - 1
          PrintThisField(3) = False
      End If
      If Trim(Pd.Home_Phone) = "" Then
          NumItems = NumItems - 1
          PrintThisField(4) = False
      End If
      If Trim(Pd.Bus_Phone) = "" Then
          NumItems = NumItems - 1
          PrintThisField(5) = False
      End If
      If Trim(Pd.Pager) = "" Then
          NumItems = NumItems - 1
          PrintThisField(6) = False
      End If
      If Trim(Pd.Fax) = "" Then
          NumItems = NumItems - 1
          PrintThisField(7) = False
      End If
      If Trim(Pd.E_Mail) = "" Then
          NumItems = NumItems - 1
          PrintThisField(8) = False
      End If
      If Trim(Pd.WebPage) = "" Then
          NumItems = NumItems - 1
          PrintThisField(9) = False
      End If
      If Trim(Pd.Notes) = "" Then
          NumItems = NumItems - 2
          PrintThisField(10) = False
      End If
   End If

'--------------------------------------------------------------------------
'Print Record
'--------------------------------------------------------------------------
AddMargin nPrintMargin
If PrintThisField(0) Then lblBody = lblBody + UserField(0) + ": " + Trim$(Pd.Company) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(1) Then lblBody = lblBody + UserField(1) + ": " + Trim$(Pd.AName) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(2) Then lblBody = lblBody + UserField(2) + ": " + Trim$(Pd.Address) + vbCrLf
If PrintThisField(3) Then
  AddMargin nPrintMargin
  lblBody = lblBody + UserField(3) + ": " + RTrim$(Pd.City) + "   " + UserField(4) + ": " + RTrim$(Pd.State) + "   " + UserField(5) + ": " + RTrim$(Pd.Zip_Code) + vbCrLf
End If
AddMargin nPrintMargin
If PrintThisField(4) Then lblBody = lblBody + UserField(6) + ": " + Trim$(Pd.Home_Phone) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(5) Then lblBody = lblBody + UserField(7) + ": " + Trim$(Pd.Bus_Phone) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(6) Then lblBody = lblBody + UserField(8) + ": " + Trim$(Pd.Pager) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(7) Then lblBody = lblBody + UserField(9) + ": " + Trim$(Pd.Fax) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(8) Then lblBody = lblBody + UserField(10) + ": " + Trim$(Pd.E_Mail) + vbCrLf
AddMargin nPrintMargin
If PrintThisField(9) Then lblBody = lblBody + UserField(11) + ": " + Trim$(Pd.WebPage) + vbCrLf

End Sub

Private Sub optColumn_Click(Index As Integer)
    GenerateScreen
End Sub

Private Sub tmrFirstShow_Timer()
    ' This timer is used since the dialog needs to become
    ' visible to work around a VB problem that caused
    ' the column calculations not to work correctly
    tmrFirstShow.Enabled = False
    GenerateScreen
    
    
End Sub
