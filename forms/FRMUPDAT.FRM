VERSION 5.00
Begin VB.Form frmUpdateCodes 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Zip Codes"
   ClientHeight    =   1155
   ClientLeft      =   1830
   ClientTop       =   3615
   ClientWidth     =   6570
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
   Icon            =   "FRMUPDAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1155
   ScaleWidth      =   6570
   Begin VB.TextBox txtChange 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "30097"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
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
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "30155"
      Top             =   240
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1095
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
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Zip Code:"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
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
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmUpdateCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
If Trim$(txtCode) = Trim$(txtChange) Then
   MsgBox "The specified codes are identical", MB_ICONEXCLAMATION, SPROGRAMNAME
ElseIf Caption = "Update Zip Codes" Then
   UpdateZipCodes
Else
   UpdateAreaCodes
End If
End Sub

Private Function GetAreaCode(PhoneNumber As String) As String
 ' Return the area code of the phone number regardless of
 ' the format (e.g. (123)456-7890 and 123-456-7890 would
 ' both return 123)

 Dim NewNumber  As String, temp As String
 Dim i As Integer
 PhoneNumber = Trim$(PhoneNumber)  'Remove leading and trailing spaces
 
 ' Remove all charactors other than digits
 NewNumber = ""
 For i = 1 To Len(PhoneNumber)
      temp = Mid$(PhoneNumber, i, 1)
      If Asc(temp) >= Asc("0") And Asc(temp) <= Asc("9") Then NewNumber = NewNumber & temp
 Next i
 
 ' Get the first three numbers, they are most likly the area code
 If Len(NewNumber) >= 9 Then
    GetAreaCode = Left$(NewNumber, 3)
 ElseIf Len(NewNumber) <= 7 Then  'The number doesn't have an area code
    GetAreaCode = ""
 End If
End Function

Private Function NewAreaCode(ByVal OldArea As String, ByVal NewArea As String, ByVal Phone As String) As String
' Change a phone numbers area code
' e.g if the phone number is (123)456-7890 and the new area
' code is 321 then (321)456-7890 whould be returned
'
' Supported formats are (123)456-7890, 123-456-7890, and 123 456-7890
' other formats are left unaltered

Dim s As Integer, e As Integer, i As Integer, Flag As Integer
Dim OrPhone As String
On Error GoTo PhoneError

OldArea = Trim$(OldArea)
NewArea = Trim$(NewArea)
Phone = Trim$(Phone)
OrPhone = Phone
s = InStr(Phone, "(")
e = InStr(Phone, ")")

Flag = 0
For i = 1 To Len(Phone)     'Get number of dashes
    If Mid$(Phone, i, 1) = "-" Then Flag = Flag + 1
Next i

If s <> 0 Then    'The phone number is the (123)456-7890 format
   If Trim$(Mid$(Phone, s + 1, e - (s + 1))) = OldArea Then
       Phone = Left$(Phone, s) & NewArea & right(Phone, Len(Phone) - (e - 1))
       NewAreaCode = Phone
   End If
ElseIf InStr(Phone, " ") <> 0 Then   '123 456-7890 format
   If Trim$(Left$(Phone, InStr(Phone, " "))) = OldArea Then
       Phone = NewArea & " " & right$(Phone, Len(Phone) - InStr(Phone, " "))
       NewAreaCode = Phone
   End If
ElseIf Flag <= 2 Then
   Do While InStr(Phone, " ") <> 0    'Remove spaces
       Phone = Left$(Phone, InStr(Phone, " ") - 1) & right$(Phone, Len(Phone) - InStr(Phone, " "))
   Loop
   e = InStr(Phone, "-")
   If Trim$(Mid$(Phone, e - 3, 3)) = OldArea Then
       Phone = Left$(Phone, e - 4) & NewArea & "-" & right$(Phone, Len(Phone) - e)
       NewAreaCode = Phone
   End If
Else
   NewAreaCode = OrPhone
End If
Exit Function

PhoneError:
  NewAreaCode = OrPhone
  Exit Function
Resume
End Function

Private Sub txtChange_Change()
 If Trim$(txtCode) = "" Then
    cmdOK.Enabled = False
 Else
    cmdOK.Enabled = True
 End If
End Sub

Private Sub txtCode_Change()
 If Trim$(txtCode) = "" Then
    cmdOK.Enabled = False
 Else
    cmdOK.Enabled = True
 End If
End Sub

Private Sub UpdateAreaCodes()
Dim msg As String, Phone1 As String, Phone2 As String, Phone3 As String, Phone4 As String
Dim ACode As String, NewACode As String, x As Integer, i As Integer
Dim Flag As Integer
ACode = Trim$(txtCode)
NewACode = Trim$(txtChange)

For i = 2 To 101
   GetPD FreeNum, i, True
   Phone1 = GetAreaCode(Pd.Home_Phone)
   Phone2 = GetAreaCode(Pd.Bus_Phone)
   Phone3 = GetAreaCode(Pd.Pager)
   Phone4 = GetAreaCode(Pd.Fax)
   If Phone1 = ACode Then
      Flag = True
      msg = "Phone: " & Pd.Home_Phone
      msg = msg & crlf & "Update area code to " & NewACode
      x = MsgBox(msg, MB_YESNOCANCEL + MB_ICONQUESTION, "Verify Update - " & Trim$(Pd.Company) & " " & Trim$(Pd.AName))
      If x = IDYES Then
         Pd.Home_Phone = NewAreaCode(ACode, NewACode, Pd.Home_Phone) 'Replace Area Code
         PutPD FreeNum, i, True
      ElseIf x = IDCANCEL Then
         Exit Sub
      End If
   End If
   
   If Phone2 = ACode Then
      Flag = True
      msg = "Phone: " & Pd.Bus_Phone
      msg = msg & crlf & "Update area code to " & NewACode
      x = MsgBox(msg, MB_YESNOCANCEL + MB_ICONQUESTION, "Verify Update - " & Trim$(Pd.Company) & " " & Trim$(Pd.AName))
      If x = IDYES Then
         Pd.Bus_Phone = NewAreaCode(ACode, NewACode, Pd.Bus_Phone) 'Replace Area Code
         PutPD FreeNum, i, True
      ElseIf x = IDCANCEL Then
         Exit Sub
      End If
   End If

   If Phone3 = ACode Then
      Flag = True
      msg = "Pager: " & Pd.Home_Phone
      msg = msg & crlf & "Update area code to " & NewACode
      x = MsgBox(msg, MB_YESNOCANCEL + MB_ICONQUESTION, "Verify Update - " & Trim$(Pd.Company) & " " & Trim$(Pd.AName))
      If x = IDYES Then
         Pd.Pager = NewAreaCode(ACode, NewACode, Pd.Pager)
         PutPD FreeNum, i, True
      ElseIf x = IDCANCEL Then
         Exit Sub
      End If
   End If

   If Phone4 = ACode Then
      Flag = True
      msg = "Fax: " & Pd.Home_Phone
      msg = msg & crlf & "Update area code to " & NewACode
      x = MsgBox(msg, MB_YESNOCANCEL + MB_ICONQUESTION, "Verify Update - " & Trim$(Pd.Company) & " " & Trim$(Pd.AName))
      If x = IDYES Then
         Pd.Fax = NewAreaCode(ACode, NewACode, Pd.Fax) 'Replace Area Code
         PutPD FreeNum, i, True
      ElseIf x = IDCANCEL Then
         Exit Sub
      End If
   End If
Next i
If Flag Then
   MsgBox "Search complete", MB_ICONEXCLAMATION, SPROGRAMNAME
   Unload Me
Else
   MsgBox "Cannot find any items with the specified area code", MB_ICONEXCLAMATION, SPROGRAMNAME
End If
End Sub

Private Sub UpdateZipCodes()
Dim msg As String, Zip1 As String, Zip2 As String, i As Integer
Dim ZCode As String, NewZCode As String, x As Integer
Dim Flag As Integer
ZCode = Trim$(txtCode)
NewZCode = Trim$(txtChange)

For i = 2 To 101
   GetPD FreeNum, i, True
   Zip1 = Trim$(Pd.Zip_Code)
   If Zip1 = ZCode Then
      Flag = True
      msg = "Address:  " & Trim$(Pd.Address) & crlf
      msg = msg & "City:  " & Trim$(Pd.City) & "  State:  " & Trim$(Pd.State) & "  Zip:  " & Zip1 & crlf
      msg = msg & crlf & "Update zip code to " & NewZCode
      x = MsgBox(msg, MB_YESNOCANCEL + MB_ICONQUESTION, "Verify Update - " & Trim$(Pd.Company) & " " & Trim$(Pd.AName))
      If x = IDYES Then
         Pd.Zip_Code = NewZCode
         PutPD FreeNum, i, True
      ElseIf x = IDCANCEL Then
         Exit Sub
      End If
   End If
Next i
If Flag Then
   MsgBox "Search complete", MB_ICONEXCLAMATION, SPROGRAMNAME
   Unload Me
Else
   MsgBox "Cannot find any items with the specified zip code", MB_ICONEXCLAMATION, SPROGRAMNAME
End If
End Sub

