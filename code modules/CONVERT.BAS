Attribute VB_Name = "CONVERT"
Option Explicit
'Original Personal Directory Manager 97 type block
Type OldPdType
     OLast_Name As String * 15
    OFirst_Name As String * 15
       OInitial As String * 1
       OAddress As String * 40
          OCity As String * 15
         OState As String * 2
      OZip_Code As String * 10
   OBus_Address As String * 40
      OBus_City As String * 15
     OBus_State As String * 2
  OBus_Zip_Code As String * 10
    OHome_Phone As String * 13
     OBus_Phone As String * 13
         OPager As String * 13
           OFax As String * 13
        OE_Mail As String * 40
End Type

Type PDtype
     CLast_Name As String * 15
    CFirst_Name As String * 15
       cInitial As String * 1
       cAddress As String * 40
          CCity As String * 15
         CState As String * 2
      CZip_Code As String * 10
   cBus_Address As String * 40
      CBus_City As String * 15
     CBus_State As String * 2
  CBus_Zip_Code As String * 10
    cHome_Phone As String * 13
     cBus_Phone As String * 20
         cPager As String * 13
           cFax As String * 13
        CE_Mail As String * 100
End Type


Global PD20 As PDtype
Global Pd10 As OldPdType

Global Const PD10Size = 257
Global Const PD20Size = 324

Global FolderFormat As Integer
'--------------------------------------------------------

Function ConvertFolder(File As String) As Integer
'---------------------------------------------------------------
'Converts old PDM 97 r3 file format to BDM 1.0 format
'---------------------------------------------------------------


Dim ConFile As String, i As Integer, FN1 As Integer, Bus As Integer, Home As Integer
Dim HAdd As String, BAdd As String, FT As String, ConSize As Integer

Screen.MousePointer = 11

'Confile = Left$(File, Len(File) - 4) & ".FOB"
'If Dir$(ConFile) <> "" Then
'   Screen.MousePointer = 0
'   x% = MsgBox("The File " & ConFile & " Already exists" & CRLF & CRLF & "Replace File?", MB_ICONEXCLAMATION + MB_YESNO, "File Exists")
'   If x% = IDNO Then Exit Function
'End If

Start:

ConFile = GetFile("Name for Converted Folder", 2)
ConFile = UCase$(Trim$(ConFile))
File = UCase$(Trim$(File))

If File = ConFile Then
   MsgBox "The converted folder cannot have the same filename as the original folder", MB_ICONEXCLAMATION, SPROGRAMNAME
   GoTo Start
ElseIf ConFile = "" Then
   ConvertFolder = False
   Screen.MousePointer = 0
   Exit Function
End If

If FolderFormat = 1 Then
   ConSize = PD10Size
ElseIf FolderFormat = -2 Then
   ConSize = PD20Size
Else
   ConvertFolder = False
   Screen.MousePointer = 0
   Exit Function
End If

ConvertFolder = False

Load frmSort
frmSort.Label1 = "Converting Folder"
frmSort.Image1.Picture = frmSort.Image2.Picture
frmSort.Show
If FolderOpen Then CloseFolder

DoEvents

FN1 = FreeFile
Open File For Random As FN1 Len = ConSize

FreeNum = FreeFile
Open ConFile For Random As FreeNum Len = PD_Size

GetPDM FN1, 1

If FolderFormat = 1 Then
  FT = Trim$(PD20.CFirst_Name)
Else
  FT = Trim$(PD20.CE_Mail)
End If

ResetPD
Pd.Company = "Personal Directory Manager 1.0"
Pd.WebPage = ENCRIPT(FT)

Put #FreeNum, 1, Pd

For i = 2 To 101
   ConvertAddress = 0
   
   GetPDM FN1, i
 If Trim$(PD20.CLast_Name) <> "" Or Trim$(PD20.CFirst_Name) <> "" Then
   HAdd = Trim$(PD20.cAddress) & Trim$(PD20.CCity) & Trim$(PD20.CState) & Trim$(PD20.CZip_Code)
   HAdd = Trim$(HAdd)
   BAdd = Trim$(PD20.cBus_Address) & Trim$(PD20.CBus_City) & Trim$(PD20.CBus_State) & Trim$(PD20.CBus_Zip_Code)
   BAdd = Trim$(BAdd)
   
   If Trim$(BAdd) = "" Then Bus = False Else Bus = True
   If Trim$(HAdd) = "" Then Home = False Else Home = True
   
   If Bus And Home Then
       frmSort.Hide
       frmWhichAddress.Show 1
       frmSort.Show
       frmSort.Refresh
       If ConvertAddress = 0 Then
           Home = True
           Bus = False
       ElseIf ConvertAddress = 1 Then
           Home = False
           Bus = True
       ElseIf ConvertAddress = -1 Then   'User canceled convert
           Close #FN1, #FreeNum
           ConvertFolder = False
           Unload frmSort
           Exit Function
       End If
   End If
   
   If Home Then
      Pd.Address = PD20.cAddress
      Pd.City = PD20.CCity
      Pd.State = PD20.CState
      Pd.Zip_Code = PD20.CZip_Code
   ElseIf Bus Then
      Pd.Address = PD20.cBus_Address
      Pd.City = PD20.CBus_City
      Pd.State = PD20.CBus_State
      Pd.Zip_Code = PD20.CBus_Zip_Code
   End If
   
   Pd.Company = " "
   Pd.AName = Trim$(PD20.CFirst_Name) & " " & Trim$(PD20.CLast_Name)
   Pd.Home_Phone = Trim$(PD20.cHome_Phone) & " "
   Pd.Bus_Phone = Trim$(PD20.cBus_Phone) & " "
   Pd.Fax = Trim$(PD20.cFax) & " "
   Pd.E_Mail = Trim$(PD20.CE_Mail) & " "
   Pd.WebPage = " "
 End If

   PutPD FreeNum, i, True
   ResetPD
Next i

ResetPD
LoadIniFields
SaveFields
   

Unload frmSort


Close #FN1, #FreeNum

File = ConFile
ConvertFolder = True
Screen.MousePointer = 0
End Function

Sub GetPDM(FileNum As Integer, CI As Integer)
 If FolderFormat = 1 Then
   Get #FileNum, CI, Pd10
     PD20.CLast_Name = Pd10.OLast_Name
     PD20.CFirst_Name = Pd10.OFirst_Name
     PD20.cInitial = Pd10.OInitial
     PD20.cAddress = Pd10.OAddress
     PD20.CCity = Pd10.OCity
     PD20.CState = Pd10.OState
     PD20.CZip_Code = Pd10.OZip_Code
     PD20.cBus_Address = Pd10.OBus_Address
     PD20.CBus_City = Pd10.OBus_City
     PD20.CBus_State = Pd10.OBus_State
     PD20.CBus_Zip_Code = Pd10.OBus_Zip_Code
     PD20.cHome_Phone = Pd10.OHome_Phone
     PD20.cBus_Phone = Pd10.OBus_Phone
     PD20.cPager = Pd10.OPager
     PD20.cFax = Pd10.OFax
     PD20.CE_Mail = Pd10.OE_Mail
  ElseIf FolderFormat = -2 Then
     Get #FileNum, CI, PD20
  End If
End Sub

