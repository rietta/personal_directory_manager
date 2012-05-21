Attribute VB_Name = "BDMAN1"
Global StartPath As String
Global NoteResize As Integer
Global savedrive As String

Public SplashScreenOn As Integer

Type UserInfoType
    UName As String
    UCompany As String
End Type
Global UserInformation As UserInfoType
Global Const prog = SPROGRAMNAME


Sub BytesUsed()
 On Error Resume Next
 Dim size As String
 
 If frmMain.lblFolder = "<no binder is currently open>" Then
    frmMain.lblStatus = "Create or Open a New File"
 Else
    size = Trim$(Str$(LOF(FreeNum) / 1000)) & " kb"
    frmMain.lblStatus = "Item " & Format(CurrentIndex - 1) & " of " & Format(ItemCount) & " total items in binder using " & Format(ItemCount * PD_Size) & " bytes of " & size
 End If
End Sub

Sub ChangeControlState(m As Integer)
    ' Activate or deactivate the different options availible to the user.

    ' Several menu items are only visible when a binder is open
    frmMain.mnuFileClose.Visible = m
    frmMain.mnuFilePrint.Visible = m
    frmMain.mnuFileBackup.Visible = m
    'frmMain.mnuFRename.Visible = m
    frmMain.mnuExportSAT.Visible = m
    frmMain.mnuEdit.Visible = m
    frmMain.MnuOUpdate.Visible = m
    frmMain.mnuOChangeDescription.Visible = m
    frmMain.mnuODefaultFolder.Visible = m
    frmMain.imgEditSortButton.Visible = m
    frmMain.imgBookmarkButton.Visible = m
    
    frmMain.mnuOChangeActiveTemplate.Visible = m
    frmMain.cmdOEMailMailingList.Visible = m

    ' Hide or show separator bars
    'frmMain.FB2.Visible = m
    frmMain.FB4.Visible = m
    frmMain.mnuOSB2.Visible = m
    frmMain.mnuOPref.Visible = m
    frmMain.mnuOSB3.Visible = m
    frmMain.mnuOSB4.Visible = m

    ' Enable or disable the toolbar and alph tabs
    frmMain.imgAddButton.Visible = m
    frmMain.imgEditButton.Visible = m
    frmMain.imgDeleteButton.Visible = m
    frmMain.imgViewButton.Visible = m
    frmMain.imgPrintButton.Visible = m
    frmMain.imgFindButton.Visible = m
    
    frmMain.lblAddButton.Visible = m
    frmMain.lblEditButton.Visible = m
    frmMain.lblDeleteButton.Visible = m
    frmMain.lblViewButton.Visible = m
    frmMain.lblPrintButton.Visible = m
    frmMain.lblFindButton.Visible = m

    frmMain.imgPlaceBookmark.Visible = m
    frmMain.imgListUp.Visible = m
    frmMain.imgListDown.Visible = m


    ' Reset the main window's title
    If m = False Then frmMain.Caption = prog

    ' Enable or Disable the print buffer icon
    If m = False Then
        frmMain.imgPrintBuffer.Picture = SplashScreen.imgPrintBufferDisabled.Picture
        frmMain.imgPrintBuffer.Enabled = False
        frmMain.imgRecycle.Picture = SplashScreen.imgRecycleDis.Picture
        frmMain.imgRecycle.Enabled = False
    Else
        frmMain.imgPrintBuffer.Picture = SplashScreen.imgPrintBufferEnabled.Picture
        frmMain.imgPrintBuffer.Enabled = True
        frmMain.imgRecycle.Picture = SplashScreen.imgRecycleEn.Picture
        frmMain.imgRecycle.Enabled = True
    End If
End Sub

Function CheckChar(CharCode As Integer) As Integer
'Check for invalid charictors
Select Case Chr$(CharCode)
    Case ":": CheckChar = 0
    Case "&": CheckChar = 0
    Case Chr$(13), Chr$(10), crlf: CheckChar = 0
    Case Else: CheckChar = CharCode
End Select
End Function

Function CheckFileFormat(Header As String) As Integer
 ' This funtion checks a files header to determine if it is a Business
 ' Directory manager folder or not.

 Header = Trim$(Header)
 If Left$(Header, 5) = "PDM20" Then   'Personal Directory Manager 2.0 binder
     CheckFileFormat = -2
 ElseIf Left$(Header, 3) = "PDM" Then 'Personal Directory Manager 1.0 binder
     CheckFileFormat = 1
 Else
     Select Case Trim$(Header)  'Check for various versions of BDM
       Case Encript("Personal Directory Manager 1.0")
             CheckFileFormat = 2
       Case "&ER&BDM20: Business Director Manager"
             CheckFileFormat = 2
       Case "Personal Directory Manager 1.0"
             CheckFileFormat = 2
       Case Else
             CheckFileFormat = 0  'Invalid File Format
     End Select
 End If
End Function

Function CheckString(ByVal StringToCheck As String, StringToFind As String) As Integer
Dim i As Integer, temp As String
For i = 1 To Len(StringToCheck)
temp = temp & Mid$(StringToCheck, i, 1)
If Mid$(StringToCheck, i, 1) = " " Or i = Len(StringToCheck) Then
    If Trim$(temp) = StringToFind Then
        CheckString = True
        Exit Function
    End If
    temp = ""
End If
Next i
CheckString = False
End Function
' Modified: 2-11-1999
Sub CloseFolder()
    On Error Resume Next
    
    Close #FreeNum
    
    CloseAllWindows
    frmMain.lblFolder = "<no binder is currently open>"
    
    FreeNum = 0
    ItemCount = 0
    BytesUsed
    
    ChangeControlState False
    
    frmMain.imgBinderHeader.Picture = frmMain.Image1.Picture
    FolderOpen = False
    
    Unload frmbookmarks
    If RecyclingBinOpen Then Unload frmRecycle
    
    LoadIniFields
    frmMain.flexItems.Clear
    frmMain.flexItems.Rows = 1
    InitializeGrid frmMain.flexItems
    
    App.Title = SPROGRAMNAME & " " & Format(nNumericVersion)
End Sub

' Delete an Item
' Modified on 2-16-1999 to sport a more informative dialog box
' instead of using the MsgBox function
Sub DeleteItem()
    Dim i As Integer, Flag As Integer, Bkm As Integer
    Dim itemName As String
           
    GetPD FreeNum, CurrentIndex, True
    
    itemName = Pd.Company & ", " & Pd.AName
    
    Load DeleteItemConfirmationDialog
    With DeleteItemConfirmationDialog
        .lblQuestion = "You have choosen to delete " & itemName & " from " & OpenFolderName & "." & vbCrLf & vbCrLf & "Which of the following do you want to do?"
        .Show 1, frmMain
    End With
    If nDialogButton = CANCELBUTTON Then Exit Sub
    
    Bkm = Pd.Bookmark
    Screen.MousePointer = 11
    If nDialogButton = WASTEBUTTON Then RecycleIt  'move item to waste bin
    Flag = 0
    For i = CurrentIndex To 101    'Shift Items following the deleted item up
        Get FreeNum, i, Pd
        If i < 101 Then
            Get FreeNum, i + 1, Pd
            Put FreeNum, i, Pd
        End If
    Next i
    ResetPD
    Put FreeNum, 101, Pd
    CurrentIndex = 2
    
    LoadList
    If Bkm And BookmarksOpen Then LoadBookmarks
    If RecyclingBinOpen Then LoadRecyclingBin
    Screen.MousePointer = 0
End Sub

Sub DoTitle(lbl As Label)
Dim msg As String
msg = UserField(0)
If Trim$(msg) <> "" Then msg = msg & ":" & Space$(25)
If Trim$(UserField(1)) <> "" Then msg = msg & UserField(1) & ":"
lbl.Caption = msg
End Sub

Sub EditItem()
    Si = "Modify Item"

    GetPD FreeNum, CurrentIndex, True
    bm% = Pd.Bookmark
    EditWindow.Show 1
    
    If RTrim$(Pd.AName) <> "" Or RTrim$(Pd.Company) <> "" Then
        Pd.Bookmark = bm%
        PutPD FreeNum, CurrentIndex, True
    End If
    LoadList
End Sub

Sub EraseHighlightedButton(win As Integer, bttn As Double)
    If win = 0 Then
        Select Case bttn
            Case 1.1: Erase3DBorderOnPicBox frmMain.imgFileNewButton, 1, 0, False, frmMain.picToolbar, frmMain.picToolbar.BackColor
            Case 1.2: Erase3DBorderOnPicBox frmMain.imgFileOpenButton, 1, 0, False, frmMain.picToolbar, frmMain.picToolbar.BackColor
            Case 1.3: Erase3DBorderOnPicBox frmMain.imgEditSortButton, 1, 0, False, frmMain.picToolbar, frmMain.picToolbar.BackColor
            Case 1.4: Erase3DBorderOnPicBox frmMain.imgBookmarkButton, 1, 0, False, frmMain.picToolbar, frmMain.picToolbar.BackColor
            
            Case 4.1: Erase3DBorderOnPicBox frmMain.imgListUp, 1, 0, False, frmMain.picListTools, frmMain.picToolbar.BackColor
            Case 4.2: Erase3DBorderOnPicBox frmMain.imgListDown, 1, 0, False, frmMain.picListTools, frmMain.picToolbar.BackColor
            Case 4.3: Erase3DBorderOnPicBox frmMain.imgPlaceBookmark, 1, 0, False, frmMain.picListTools, frmMain.picToolbar.BackColor
        End Select
    ElseIf win = 1 Then
        Select Case bttn
            Case 1.1: Erase3DBorderOnPicBox frmMDI.imgFileNewButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            Case 1.2: Erase3DBorderOnPicBox frmMDI.imgFileOpenButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            Case 1.3: Erase3DBorderOnPicBox frmMDI.imgSaveButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            Case 1.4: Erase3DBorderOnPicBox frmMDI.imgPrintButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            
            Case 2.1: Erase3DBorderOnPicBox frmMDI.imgCutButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            Case 2.2: Erase3DBorderOnPicBox frmMDI.imgCopyButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            Case 2.3: Erase3DBorderOnPicBox frmMDI.imgPasteButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor
            
            Case 3.1: Erase3DBorderOnPicBox frmMDI.imgMainWinButton, 1, 0, False, frmMDI.picToolbar, frmMain.picToolbar.BackColor

        End Select
    End If
End Sub

Sub FolderName(DN$)
   ResetPD
   Pd.Company = "&ER&BDM20: Business Director Manager"
   Pd.WebPage = Encript(DN$)
   Put #FreeNum, 1, Pd
End Sub

Function Generate8CharFileName(ByVal File As String)
' ----------------------------------------------------
' Generate a suggested file name from the description
' ----------------------------------------------------
Dim temp As String, i As Integer
File = Trim$(File)
If Len(File) > 5 Then File = Left$(File, 5)

temp = ""
For i = 1 To Len(File)
    Select Case Mid$(File, i, 1)
         Case " ", "*", "?", "/", "\", "|", ".", ",", ";", ":", "+", "=", "[", "]", "(", ")", "&", "^", "<", ">", """"   'Invalid charactor
         Case Else
             temp = temp & Mid$(File, i, 1)
    End Select
Next i
File = temp

' Search for free ~1 extention (like Windows 95)
If right$(dpath$, 1) <> "\" Then dpath$ = dpath$ & "\"
For i = 1 To 99
  If Dir$(dpath$ & File & "~" & Trim$(Str$(i)) & ".FOB") = "" Then
        File = File & "~" & Trim$(Str$(i)) & ".fob"
        File = LCase$(File)
        Generate8CharFileName = File
        Exit Function
  End If
Next i
Generate8CharFileName = ""
End Function

Function GetDefaultField(i As Integer) As String
'Return the default field title for the field index
Dim DF As String
Select Case i
    Case 0: DF = "Company"
    Case 1: DF = "Name"
    Case 2: DF = "Address"
    Case 3: DF = "City"
    Case 4: DF = "State"
    Case 5: DF = "Zip"
    Case 6: DF = "Home Phone"
    Case 7: DF = "Bus. Phone"
    Case 8: DF = "Pager"
    Case 9: DF = "Fax"
    Case 10: DF = "E-Mail"
    Case 11: DF = "Web Page (URL)"
    Case 12: DF = "Notes"
End Select
GetDefaultField = DF
End Function

Function GetFile(Title As String, Act As Integer) As String
    Dim retVal, OpenFileName As String
    On Error Resume Next
    If Title <> "" Then frmMain.dialog.DialogTitle = Title
    If Not txt Then
       frmMain.dialog.Filter = "All binders|*.FOB;*.FOL|BDM Binders (*.FOB)|*.FOB|PDM Binders(*.FOL)|*.FOL|All Files(*.*)|*.*"
    Else
       '  User selected: save binder as text... from the main screen
       frmMain.dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    End If
    frmMain.dialog.FilterIndex = 1
    frmMain.dialog.FileName = ""
    frmMain.dialog.FLAGS = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
    If Not txt Then
       frmMain.dialog.DefaultExt = "FOB"
    Else
       frmMain.dialog.DefaultExt = "TXT"
    End If
    frmMain.dialog.Action = Act
    If Err <> 32755 Then 'user pressed cancel
        OpenFileName = frmMain.dialog.FileName
        If OpenFileName <> "" Then GetFile = OpenFileName
    End If

End Function

Function GetFName(ByVal File As String) As String

 'Extract File name from the path string
 Dim i As Integer, LastBackSlash As Integer, temp As String
 For i = 1 To Len(File)
   temp = Mid(File, i, 1)
   If temp = "\" Then LastBackSlash = i    'Record position of last backslash in string
 Next i
 GetFName = right$(File, Len(File) - LastBackSlash)
End Function

Function GetIP()
   Load frmGenericList
   frmGenericList.Caption = "Common Internet Service Providers"
   frmGenericList.Label1 = "Please select the service to search for:"
   frmGenericList.List1.AddItem "America Online"
   frmGenericList.List1.AddItem "CompuServe"
   frmGenericList.List1.AddItem "Prodigy"
   frmGenericList.List1.AddItem "At&t WorldNet"
   frmGenericList.List1.AddItem "Mind Spring"
   frmGenericList.List1.AddItem "BellSouth.Net"
   frmGenericList.List1.AddItem "BellAtlantic.Net"
   frmGenericList.List1.AddItem "The Microsoft Network"
   frmGenericList.List1.AddItem "IBM.Net"
   frmGenericList.List1.AddItem "Negia"
   frmGenericList.List1.AddItem "Geocities"
   frmGenericList.List1.AddItem "Juno"
   frmGenericList.List1.ListIndex = 0
   
   frmGenericList.Show 1
 
   Select Case Si
        Case "America Online": GetIP = "@aol.com"
        Case "CompuServe": GetIP = "@compuserve.com"
        Case "Prodigy": GetIP = "@prodigy.com"
        Case "At&t WorldNet": GetIP = "@worldnet.att.net"
        Case "Mind Spring": GetIP = "@mindspring.com"
        Case "BellSouth.Net": GetIP = "@bellsouth.net"
        Case "BellAtlantic.Net": GetIP = "@bellatlantic.net"
        Case "The Microsoft Network": GetIP = "@msn.com"
        Case "IBM.Net": GetIP = "@ibm.net"
        Case "Negia": GetIP = "@negia.net"
        Case "Geocities": GetIP = "@geocities.com"
        Case "Juno": GetIP = "@juno.com"
        Case Else: GetIP = ""
   End Select
End Function

Sub GetPD(FileNumber As Integer, Idx As Integer, DoCript As Integer)
'Decript varibles after getting them from the file
  Get #FileNumber, Idx, Pd
  If DoCript Then
        Pd.Company = Decript(Trim$(Pd.Company))
        Pd.AName = Decript(Trim$(Pd.AName))
        Pd.Address = Decript(Trim$(Pd.Address))
        Pd.City = Decript(Trim$(Pd.City))
        Pd.State = Decript(Trim$(Pd.State))
        Pd.Zip_Code = Decript(Trim$(Pd.Zip_Code))
        Pd.Home_Phone = Decript(Trim$(Pd.Home_Phone))
        Pd.Bus_Phone = Decript(Trim$(Pd.Bus_Phone))
        Pd.Pager = Decript(Trim$(Pd.Pager))
        Pd.Fax = Decript(Trim$(Pd.Fax))
        Pd.E_Mail = Decript(Trim$(Pd.E_Mail))
        Pd.WebPage = Decript(Trim$(Pd.WebPage))
        Pd.Notes = Decript(Trim$(Pd.Notes))
  End If
End Sub

Sub LoadBookmarks()
    Dim i As Integer, ntn As Integer
    Screen.MousePointer = 11
    Si = ""
    ntn = 2
 
    If Not BookmarksOpen Then Load frmbookmarks
 
    frmbookmarks.List1.Clear
    For i = 2 To 100
        GetPD FreeNum, i, True
        If Pd.Bookmark = True Then
            frmbookmarks.List1.AddItem Trim$(Pd.Company) & Chr$(9) & Trim$(Pd.AName)
            BMIndex(ntn) = i
            ntn = ntn + 1
        End If
    Next i
 
    If ntn = 2 Then
        frmbookmarks.List1.AddItem "No Bookmarks in Binder"
        frmbookmarks.List1.Enabled = False
        frmbookmarks.cmdBookmarkView.Enabled = False
        frmbookmarks.cmdBookmarkRemove.Enabled = False
    Else
        frmbookmarks.List1.Enabled = True
        frmbookmarks.cmdBookmarkView.Enabled = True
        frmbookmarks.cmdBookmarkRemove.Enabled = True
    End If

    frmbookmarks.Show
    Screen.MousePointer = 0
    If ntn <> 2 Then frmbookmarks.List1.ListIndex = 0
 End Sub

Sub LoadFields()
Dim User1 As String, User2 As String, temp As String
Dim i As Integer, CurrentLoc As Integer
GetPD FreeNum, 102, True  'Load Field Names from file
User1 = Trim$(Pd.E_Mail)
User2 = Trim$(Pd.WebPage)
UserField(12) = Trim$(Pd.Notes)
'-----------------------------------------------
'Parse Strings and retrive field names
'-----------------------------------------------
CurrentLoc = 0
temp = ""
For i = 1 To Len(User1)
  If Mid$(User1, i, 1) = ":" Then
      UserField(CurrentLoc) = temp
      temp = ""
      CurrentLoc = CurrentLoc + 1
      If CurrentLoc = 12 Then Exit Sub
  Else
      temp = temp & Mid$(User1, i, 1)
  End If
Next i


temp = ""
For i = 1 To Len(User2)
  If Mid$(User2, i, 1) = ":" Then
      UserField(CurrentLoc) = temp
      temp = ""
      CurrentLoc = CurrentLoc + 1
      If CurrentLoc = 12 Then Exit Sub
  Else
      temp = temp & Mid$(User2, i, 1)
  End If
Next i

End Sub

Sub LoadFillData()
    GetPD FreeNum, 114, True

    DataToFill.Company = Pd.Company
    DataToFill.AName = Pd.AName
    DataToFill.Address = Pd.Address
    DataToFill.City = Pd.City
    DataToFill.State = Pd.State
    DataToFill.Zip_Code = Pd.Zip_Code
    DataToFill.Home_Phone = Pd.Home_Phone
    DataToFill.Bus_Phone = Pd.Bus_Phone
    DataToFill.Pager = Pd.Pager
    DataToFill.Fax = Pd.Fax
    DataToFill.E_Mail = Pd.E_Mail
    DataToFill.WebPage = Pd.WebPage
    DataToFill.Notes = Pd.Notes

End Sub

Sub LoadIniFields()
'Load the global field settings from BDMAN.INI
Dim temp As String, i As Integer
For i = 0 To 12
   UserField(i) = ""
Next i
Select Case CurrentFolderTemplate
    Case 0
        For i = 0 To 12
            UserField(i) = GetDefaultField(i)
        Next i
    Case 1
        For i = 0 To 12
              temp = GetMyIni("BDMAN.INI", "Fields", "Field" & Trim$(Str$(i)))
              temp = Trim$(temp)
              If temp = "" Then
                   UserField(i) = GetDefaultField(i)  'Get the default field
              Else
                   UserField(i) = temp
              End If
        Next i
    Case 2
        UserField(0) = "Description"
        UserField(11) = "Web Page (URL)"
        UserField(12) = "Notes"
End Select
End Sub

' Modified: 2-11-1999
Sub LoadList()
     'Dim j As Integer
     'frmMain.lstNames.Clear
     'For j = 2 To 100
     '   Get #FreeNum, j, Pd
     '   Pd.Company = Decript(Trim$(Pd.Company))
     '   Pd.AName = Decript(Trim$(Pd.AName))
     '   If RTrim$(Pd.AName) <> "" Or RTrim$(Pd.Company) <> "" Then
     '     frmMain.lstNames.AddItem Trim$(Pd.Company) & Chr$(9) & Trim$(Pd.AName)
     '   End If
     'Next j
     'ResetPD
     'On Error Resume Next
     'frmMain.lstNames.ListIndex = CurrentIndex - 2
     'frmMain.lstNames.SetFocus
     'BytesUsed
     'Screen.MousePointer = 0
     
     LoadItemsIntoGrid frmMain.flexItems
End Sub

Sub LoadPrefrences()
    ' Load Prefrences from HPhone string in header
    Dim PString As String
    GetPD FreeNum, 1, False
    PString = Trim$(Pd.Home_Phone)
    If PString = "" Then
        Prefrences.ShowStateButton = 0
        Prefrences.ShowISPButton = 0
        Prefrences.FillData = 0
    Else
        Prefrences.ShowStateButton = Val(Left$(PString, 1))
        Prefrences.ShowISPButton = Val(Mid$(PString, 2, 1))
        Prefrences.FillData = Val(right$(PString, 1))
    End If
End Sub

Sub LoadQuickList()
  ' Save the contents of the quick list
  Dim i As Integer, merge As Integer, Fetch As String
  For i = 0 To 9
     Fetch = Trim$(GetMyIni("BDMAN.INI", "Quick List", "QLINK" & Trim$(Str$(i))))
     If Fetch = "" Or Fetch = "&" Then
          QuickList(i).Description = ""
          QuickList(i).File = ""
     Else
          merge = InStr(Fetch, "&")
          QuickList(i).Description = Left$(Fetch, merge - 1)
          QuickList(i).File = right$(Fetch, Len(Fetch) - merge)
     End If
  Next i
End Sub

Sub LoadRecyclingBin()
'----------------------------------------------------------------------
'Load items that are in the waste bin
'
'Records 103-113
'----------------------------------------------------------------------
Dim i As Integer, Flag As Integer
Dim Comp As String, Nam As String

'----------------------------------------------------------------------
'Find first availible waste record to delete.  If the waste bin
'is full then over write the oldest record.
'----------------------------------------------------------------------
Flag = False
frmRecycle.List1.Clear
For i = 103 To 113
   Get #FreeNum, i, Pd
   Pd.Company = Decript$(Trim$(Pd.Company))
   Pd.AName = Decript$(Trim$(Pd.AName))
   Comp = Trim$(Pd.Company)
   Nam = Trim$(Pd.AName)
   If Trim$(Comp & Nam) <> "" Then
       frmRecycle.List1.AddItem Trim$(Pd.Company) & Chr$(9) & Trim$(Pd.AName)
       Flag = True
   End If
Next i
If Flag = False Then  'The waste Bin is empty.
  frmRecycle.cmdUndelete.Enabled = False
  frmRecycle.cmdView.Enabled = False
  frmRecycle.cmdDelete.Enabled = False
  frmRecycle.List1.AddItem "<The Waste Bin is Empty>"
  frmRecycle.List1.Enabled = False
  frmRecycle.lblDeletedItems = "Deleted Items:"
Else
  frmRecycle.lblDeletedItems = frmRecycle.List1.ListCount & " Deleted Items:"
End If
frmRecycle.List1.ListIndex = 0

RecyclingBinOpen = True
Screen.MousePointer = 0
End Sub

' Name: Main
' Purpose: First sub to execute
' Created: 12/96
' Modified: 2/11/1999
Sub Main()
    
    Dim CurPath As String
    Dim RegisteredPath As String
    Dim x As Integer
    Dim UseRef As Integer
  
    Screen.MousePointer = vbHourglass
    
    MarkAllWindowsClosed
    
    CurrentFolderTemplate = 1   'Set to normal template
  
    SplashScreen.Show
  
    SplashScreenOn = True
    SplashScreen.Timer1.Enabled = True
    
    ' wait a few seconds
    Do While SplashScreenOn
        DoEvents
    Loop
    
    
    'evaluate command line
    Select Case Trim$(LCase$(Command$))
        Case "/notebook":
            frmMDI.Show
            NoteBook = 2
            DoEvents
        Case "/file search"
            MsgBox "The file search feature of Personal Directory Manager has been removed in version 2.6 because it no longer serves a purpose in Windows 95/98.", vbInformation, SPROGRAMNAME
        Case Else
            frmMain.Show
            DoEvents
    End Select
    
    Unload SplashScreen
          
    CurPath = LCase$(App.Path)
    If right$(CurPath, 1) <> "\" Then CurPath = CurPath & "\"
    
    InitializeTemplateInfo
    LoadQuickList
    ShowQuickList
    
    Screen.MousePointer = vbNormal
End Sub

'   Add the items in the sourcefile to the destfile
Function MergeBindersAdd(sourcefile As String, destfile As String) As Integer
    Dim i As Integer, d As Integer, freen1 As Integer, freen2 As Integer
    Dim sCount As Integer, dCount As Integer
    
    MergeBindersAdd = False
    
    freen2 = FreeFile
    Open destfile For Random As #freen2 Len = PD_Size
    dCount = NumberOfItems(freen2)
    If dCount > 99 Then
        MsgBox destfile & " is full and no more items can be added to it.", MB_ICONEXCLAMATION, "Binder Full"
        Close #freen2
        Exit Function
    End If

    freen1 = FreeFile
    Open sourcefile For Random As #freen1 Len = PD_Size
    sCount = NumberOfItems(freen1)

    If sCount > (99 - dCount) Then
        i = MsgBox(sourcefile & " has more items in it than " & destfile & " has free space.  Do you wish to add all the items can fit into the latter file?", MB_ICONQUESTION + MB_YESNO, "More Items than Space")
        If i = IDNO Then
            Close #freen1, freen2
            Exit Function
        End If
    End If

    d = dCount + 2
    For i = 2 To sCount + 2
        GetPD freen1, i, False
        d = d + 1
        If d >= 100 Then Exit For
        PutPD freen2, d, False
    Next i
    Close #freen1, freen2
    
    MsgBox sourcefile & " has been successfully merged into " & destfile, MN_ICONINFORMATION, "Merge Successful"
    MergeBindersAdd = True
End Function

Function NameFolder(WinTitle As String, Templates As Integer) As String
'Get the users input on what the new name of the binder should be
Load BinderDescriptionDialog
BinderDescriptionDialog.Caption = WinTitle
BinderDescriptionDialog.txtFolderName.MaxLength = 100
BinderDescriptionDialog.txtFolderName.text = OpenFolderName
BinderDescriptionDialog.txtFolderName.SelStart = 0
BinderDescriptionDialog.txtFolderName.SelLength = Len(OpenFolderName)
BinderDescriptionDialog.lblTemplate.Visible = Templates
BinderDescriptionDialog.cboTemplates.Visible = Templates
BinderDescriptionDialog.Show 1
NameFolder = Trim$(Si$)
End Function

Sub NewCaption(FileName As String)
 CurrentCaption = " " & UCase$(GetFName(FileName))
 CurrentCaption = prog & " - " & CurrentCaption
 frmMain.Caption = CurrentCaption
 App.Title = CurrentCaption
End Sub

Function NoBlankFields()
' open item in notebook
   GetPD FreeNum, CurrentIndex, True
   Dim msg As String
   
   msg = ""
   
   If Trim$(Pd.Company) <> "" Then msg = UserField(0) & ": " & RTrim$(Pd.Company)
   If Trim$(Pd.AName) <> "" Then msg = msg & crlf & UserField(1) & ": " & RTrim$(Pd.AName)
   If Trim$(Pd.Address) <> "" Then msg = msg & crlf & UserField(2) & ": " & RTrim$(Pd.Address)
   If Trim$(Pd.City) <> "" Then msg = msg & crlf & UserField(3) & ": " & RTrim$(Pd.City)
   If Trim$(Pd.State) <> "" Then msg = msg & crlf & UserField(4) & ": " & RTrim$(Pd.State)
   If Trim$(Pd.Zip_Code) <> "" Then msg = msg & crlf & UserField(5) & ": " & RTrim$(Pd.Zip_Code)
   If Trim$(Pd.Home_Phone) <> "" Then msg = msg & crlf & UserField(6) & ": " & RTrim$(Pd.Home_Phone)
   If Trim$(Pd.Bus_Phone) <> "" Then msg = msg & crlf & UserField(7) & ": " & RTrim$(Pd.Bus_Phone)
   If Trim$(Pd.Pager) <> "" Then msg = msg & crlf & UserField(8) & ": " & RTrim$(Pd.Pager)
   If Trim$(Pd.Fax) <> "" Then msg = msg & crlf & UserField(9) & ": " & RTrim$(Pd.Fax)
   If Trim$(Pd.E_Mail) <> "" Then msg = msg & crlf & UserField(10) & ": " & RTrim$(Pd.E_Mail)
   If Trim$(Pd.WebPage) <> "" Then msg = msg & crlf & UserField(11) & ": " & RTrim$(Pd.WebPage)
   If Trim$(Pd.Notes) <> "" Then msg = msg & crlf & UserField(12) & ": " & RTrim$(Pd.Notes)

   NoBlankFields = msg
End Function

Function NumberOfItems(freen As Integer) As Integer
    Dim i As Integer, n As Integer

    n = 0
    For i = 2 To 100
        GetPD freen, i, False
        If Trim$(Pd.Company) & Trim$(Pd.AName) <> "" Then
            n = n + 1
        Else
            Exit For
        End If
    Next i
    NumberOfItems = n
End Function

Sub OpenFolder(FileName As String)
On Error GoTo OpenError
Dim x As Integer, FileNum As Integer, CheckHeader As String, i As Integer
'--------------------------------------------------------------------
'Check for a valid file format

'Personal Directory Manager 97 r3 Folders *.FOL file header is: "PDM"

'Personal Directory Manager 97 Folders *.FOB file header is:
'"Personal Directory Manager 97 S 1.0a"

'Personal Directory Manager 2.0 file header: "PDM20"
'--------------------------------------------------------------------
    FileName = Trim$(FileName)
      
    If FileName = "" Then
        Exit Sub
    ElseIf LCase$(FileName) = LCase$(sBinderFileName) Then
        MsgBox "The file " & """" & FileName & """" & "is already open.", MB_ICONINFORMATION, prog
        Exit Sub
    End If
        
      
      FileNum = FreeFile
      
      On Error Resume Next
      
      frmMain.lblStatus = "Loading File..."
      frmMain.lblStatus.Refresh
      
      Open FileName For Input As #FileNum
      If Err Then
          MsgBox "Cannot open: " & Trim$(UCase$(FileName)), MB_ICONSTOP, SPROGRAMNAME
          Exit Sub
      End If
      CheckHeader = Input$(100, #FileNum)
      If Err = 51 Then GoTo ErrorFormat
      Close #FileNum
      FolderFormat = CheckFileFormat(CheckHeader)
      Select Case FolderFormat
          Case 0
ErrorFormat:
             MsgBox "File: " & Trim$(UCase$(FileName)) & " is an unreconized file format", MB_ICONSTOP, SPROGRAMNAME
             Exit Sub
          Case 1, -2
             x = 0
             x = MsgBox(Trim$(UCase$(FileName)) & " is a Personal Directory Manager binder.  Convert binder to Personal Directory Manager format", MB_ICONQUESTION + MB_YESNO, "Personal Directory Manager Binder")
             If x = IDYES Then
                  If FolderOpen Then CloseFolder
                  x = ConvertFolder(FileName)
                  If x = False Then
                      MsgBox "Cannot convert Binder", MB_ICONSTOP, SPROGRAMNAME
                      Exit Sub
                  End If
             Else
                  Exit Sub
             End If
          Case 2
             'Correct Format don't do anything
      End Select
      x = 0
'---------------------------------------------------------------------------
' Initialize Opening
'---------------------------------------------------------------------------
      Screen.MousePointer = 11
      If FolderOpen Then Close #FreeNum                   'Close open Binder
      FreeNum = FreeFile                                  'Get free file #
      Open FileName For Random As #FreeNum Len = PD_Size  'Open Binder File
      sBinderFileName = FileName                                   'Set New File name
                                                          'in global varible
'---------------------------------------------------------------------------
'Load Information
'---------------------------------------------------------------------------
      Get #FreeNum, 1, Pd
      frmMain.lblFolder = Decript(Trim$(Pd.WebPage))   'Place binders Long name in the Binder
                                       'name box.
      LoadList                         'Load contents into the list.
      
      LoadItemsIntoGrid frmMain.flexItems
      
      BytesUsed                        'Show Byte Information
                                       'info bar.
      ChangeControlState True          'Activate the Toolbar and Edit menu
      'Change the binder icon to visuly indicate that the binder is open
      frmMain.imgBinderHeader.Picture = frmMain.Image2.Picture
      FolderOpen = True                'Allow other parts of the program to
                                       'know that a binder is currently open.
      ResetPD                          'Reset type varible
      CurrentIndex = 2                 'Set Current File location to the
                                       'first user record.
      
      LoadFillData
      LoadPrefrences
      LoadFields
      
      ' Check to see if any fields are defined
      Flag = False
      For i = 0 To 12
          If Trim$(UserField(i)) <> "" Then
              Flag = True
              Exit For
          End If
      Next i
      
      If Not Flag Then
          MsgBox "There are no fields defined for this binder.  You must define some fields before using this binder.", MB_ICONEXCLAMATION, SPROGRAMNAME
          frmFields.Show 1
      End If
      
      InitializeGrid frmMain.flexItems
'---------------------------------------------------------------------------
'Check if this file is the default Binder
'---------------------------------------------------------------------------
      DefaultFolder$ = GetMyIni("BDMAN.INI", "Recent Files", "DefFolder")
      If UCase$(FileName) = UCase$(DefaultFolder$) Then
           frmMain.mnuODefaultFolder.Checked = True
      Else
           frmMain.mnuODefaultFolder.Checked = False
      End If
'---------------------------------------------------------------------------
'Show the main window with the open Folder
'---------------------------------------------------------------------------
      frmMain.WindowState = 0
      frmMain.Show
      frmMain.flexItems.SetFocus
      NewCaption FileName   'Redo the main forms caption to
                            'include the open file name.


      ' If other windows are open then update their contents
      If BookmarksOpen Then LoadBookmarks
      BytesUsed
Screen.MousePointer = 0
Exit Sub

OpenError:
   Beep
   MsgBox Error$, MB_ICONSTOP, SPROGRAMNAME
   Exit Sub
Resume

End Sub

Function BDM_PrinterSetup() As Integer
    Load SplashScreen
    SplashScreen.dlog.Copies = 1
    SplashScreen.dlog.PrinterDefault = True
    SplashScreen.dlog.FLAGS = PD_PRINTSETUP 'Or PD_SHOWHELP
    SplashScreen.dlog.Action = 5
    Unload SplashScreen
End Function

Sub PutPD(FileNumber As Integer, Idx As Integer, DoCript As Integer)
  'Encript varibles before puting them in the file.
  'Bookmark field is not incripted because it is an integer.
  If DoCript Then
      Pd.Company = Encript(Trim$(Pd.Company))
      Pd.AName = Encript(Trim$(Pd.AName))
      Pd.Address = Encript(Trim$(Pd.Address))
      Pd.City = Encript(Trim$(Pd.City))
      Pd.State = Encript(Trim$(Pd.State))
      Pd.Zip_Code = Encript(Trim$(Pd.Zip_Code))
      Pd.Home_Phone = Encript(Trim$(Pd.Home_Phone))
      Pd.Bus_Phone = Encript(Trim$(Pd.Bus_Phone))
      Pd.Pager = Encript(Trim$(Pd.Pager))
      Pd.Fax = Encript(Trim$(Pd.Fax))
      Pd.E_Mail = Encript(Trim$(Pd.E_Mail))
      Pd.WebPage = Encript(Trim$(Pd.WebPage))
      Pd.Notes = Encript(Trim$(Pd.Notes))
  End If
  Put #FileNumber, Idx, Pd
End Sub

Sub RecycleIt()
'----------------------------------------------------------------------
'Place the deleted data into the recycling bin.
'
'Record locations 103 - 113 are set asside for the waste bin.
'----------------------------------------------------------------------
Dim i As Integer, RecycleTo As Integer, Flag As Integer
Dim Comp As String, Nam As String
'----------------------------------------------------------------------
'Find first availible waste record to delete.  If the waste bin
'is full then over write the oldest record.
'----------------------------------------------------------------------
Flag = False
For i = 103 To 113
   GetPD FreeNum, i, True
   Comp = Trim$(Pd.Company)
   Nam = Trim$(Pd.AName)
   If (Comp = "" Or Comp = Null) And (Nam = "" Or Nam = Null) Then
       'Record is free.  So record it's location
       RecycleTo = i
       Flag = True
       Exit For
   End If
   If EOF(FreeNum) Then Exit For
Next i
If Flag = False Then
                'The Waste Bin is full so over write the oldest record.
    For i = 103 To 112           'Shift records up
       Get #FreeNum, i + 1, Pd
       Put #FreeNum, i, Pd
    Next i
    RecycleTo = 113
End If
Get #FreeNum, CurrentIndex, Pd
Put #FreeNum, RecycleTo, Pd            'Save Record


End Sub

Function ReplaceFile(File As String) As Integer
Dim x As Integer
x = MsgBox(File & " already exists" & crlf & crlf & "Replace " & File & "?", MB_YESNO + MB_ICONQUESTION, "Replace File")
If x = IDYES Then ReplaceFile = True
   
End Function

Sub ResetPD()
 Screen.MousePointer = 11
  Pd10.OLast_Name = " "
  Pd10.OFirst_Name = " "
  Pd10.OInitial = " "
  Pd10.OAddress = " "
  Pd10.OCity = " "
  Pd10.OState = " "
  Pd10.OZip_Code = " "
  Pd10.OBus_Address = " "
  Pd10.OBus_City = " "
  Pd10.OBus_State = " "
  Pd10.OBus_Zip_Code = " "
  Pd10.OHome_Phone = " "
  Pd10.OBus_Phone = " "
  Pd10.OPager = " "
  Pd10.OFax = " "
  Pd10.OE_Mail = " "
  
  PD20.CLast_Name = " "
  PD20.CFirst_Name = " "
  PD20.cInitial = " "
  PD20.cAddress = " "
  PD20.CCity = " "
  PD20.CState = " "
  PD20.CZip_Code = " "
  PD20.cBus_Address = " "
  PD20.CBus_City = " "
  PD20.CBus_State = " "
  PD20.CBus_Zip_Code = " "
  PD20.cHome_Phone = " "
  PD20.cBus_Phone = " "
  PD20.cPager = " "
  PD20.cFax = " "
  PD20.CE_Mail = " "
  
  Pd.Company = " "
  Pd.AName = " "
  Pd.Address = " "
  Pd.City = " "
  Pd.State = " "
  Pd.Zip_Code = " "
  Pd.Home_Phone = " "
  Pd.Bus_Phone = " "
  Pd.Pager = " "
  Pd.Fax = " "
  Pd.E_Mail = " "
  Pd.WebPage = " "
  Pd.Notes = " "
  Pd.Bookmark = False
  Screen.MousePointer = 0
End Sub

Sub SaveFields()
    Dim User1 As String, User2 As String
    User1 = ""
    User2 = ""
    
    User1 = User1 & UserField(0) & ":"
    User1 = User1 & UserField(1) & ":"
    User1 = User1 & UserField(2) & ":"
    User1 = User1 & UserField(3) & ":"
    User1 = User1 & UserField(4) & ":"
    User1 = User1 & UserField(5) & ":"
    
    User2 = User2 & UserField(6) & ":"
    User2 = User2 & UserField(7) & ":"
    User2 = User2 & UserField(8) & ":"
    User2 = User2 & UserField(9) & ":"
    User2 = User2 & UserField(10) & ":"
    User2 = User2 & UserField(11) & ":"
    
    Pd.E_Mail = User1
    Pd.WebPage = User2
    Pd.Notes = UserField(12)
    
    PutPD FreeNum, 102, True

End Sub

Sub SaveQuickList()
  ' Save the contents of the quick list
  Dim i As Integer
  For i = 0 To 9
     SetMyIni "BDMAN.INI", "Quick List", "QLINK" & Trim$(Str$(i)), QuickList(i).Description & "&" & QuickList(i).File
  Next i
End Sub

Sub Search()
 Dim i As Integer, ntn As Integer, SearchFor As String, Match As Integer
 Dim StringToCheck As String

 SearchFor = SearchData.searchText
 If Not SearchData.CaseSens Then SearchFor = UCase$(SearchFor)
 
 Screen.MousePointer = 11
 
 ntn = 2
 
 frmLocateIt.List1.Clear
   For i = 2 To 100
     GetPD FreeNum, i, True    ' Get Next Record
     Match = False
     Select Case SearchData.SearchMode
        Case 0  ' Company
            StringToCheck = Trim$(Pd.Company)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 1   ' Name
            StringToCheck = Trim$(Pd.AName)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 2   ' Address
            StringToCheck = Trim$(Pd.Address)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 3   ' City
            StringToCheck = Trim$(Pd.City)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 4   ' State
            StringToCheck = Trim$(Pd.State)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            If StringToCheck = SearchFor Then Match = True
        Case 5   ' Zip
            StringToCheck = Trim$(Pd.Zip_Code)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 6   ' h. phone
            StringToCheck = Trim$(Pd.Home_Phone)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 7   ' b. phone
            StringToCheck = Trim$(Pd.Bus_Phone)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 8   ' fax
            StringToCheck = Trim$(Pd.Fax)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 9   ' pager
            StringToCheck = Trim$(Pd.Pager)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 10  ' e-mail
            StringToCheck = Trim$(Pd.E_Mail)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
            If InStr(StringToCheck, SearchFor) > 0 Then Match = True ' Allow for service provider searches
        Case 11  ' web page
            StringToCheck = Trim$(Pd.WebPage)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True
        Case 12  ' notes
            StringToCheck = Trim$(Pd.Notes)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True

        Case 13  ' All Fields
            StringToCheck = Trim$(Pd.Company) & " " & Trim$(Pd.AName) & " " & Trim$(Pd.Address) & " " & Trim$(Pd.City) & " " & Trim$(Pd.State) & " " & Trim$(Pd.Zip_Code) & " " & Trim$(Pd.Home_Phone) & " " & Trim$(Pd.Bus_Phone) & " " & Trim$(Pd.Pager) & " " & Trim$(Pd.Fax) & " " & Trim$(Pd.E_Mail) & " " & Trim$(Pd.WebPage) & " " & Trim$(Pd.Notes)
            If Not SearchData.CaseSens Then StringToCheck = UCase$(StringToCheck)
            Match = CheckString(StringToCheck, SearchFor)
            If StringToCheck = SearchFor Then Match = True

     End Select

     If Match Then
         frmLocateIt.List1.AddItem Trim$(Pd.Company) & Chr$(9) & Trim$(Pd.AName)
         frmLocateIt.Refresh
         SRIndex(ntn) = i
         ntn = ntn + 1
     End If
    Next i
    
    Screen.MousePointer = 0
    If ntn = 2 Then
        frmLocateIt.List1.Clear
        frmLocateIt.List1.AddItem "No Matches Found"
        frmLocateIt.List1.Enabled = False
        frmLocateIt.cmdView.Enabled = False
    Else
        frmLocateIt.List1.ListIndex = 0
        frmLocateIt.List1.Enabled = True
        frmLocateIt.cmdView.Enabled = True
    End If

End Sub

Function SelectFont() As Integer
 Dim temp As String
 Screen.MousePointer = 11
 temp = SFont$
 Fonts.Show 1
 If SFont = "" Then
     SFont = temp
     SelectFont = False
 Else
     SelectFont = True
 End If
End Function

Sub SetTabs(frm As Form, lst As ListBox)
     Static tabs(3) As Integer
     
     Dim j As Integer, hOldWnd As Long, lbhWnd As Long, retVal As Long, r As Long
     
     Const LB_SETTABSTOPS = WM_USER + 19
     Screen.MousePointer = 11
     hOldWnd = GetFocus()   'Remember who had the focus.
     frm.Show
     lst.SetFocus          'Set the focus to the list box.
     lbhWnd = GetFocus()    'Get the handle to the list box.
     'Set up the array of defined tab stops.
     tabs(1) = 90
     tabs(2) = 1
     tabs(3) = 1
     'Send a message to the message queue.
     retVal = SendMessage(lbhWnd, LB_SETTABSTOPS, 3, tabs(1))
     'Restore the handle to whoever had it.
     r = PutFocus(hOldWnd)
End Sub

Sub ShowQuickList()
  ' Load the contents of the Quick List into the pulldown menu
  Dim i As Integer, Flag As Integer
  Flag = False
  For i = 0 To 9
    If QuickList(i).Description <> "" Then
        frmMain.mnuQuickLink(i).Caption = QuickList(i).Description & ":  " & Chr$(9) & GetFileFromPath(QuickList(i).File)
        frmMain.mnuQuickLink(i).Visible = True
        Flag = True
    Else
        frmMain.mnuQuickLink(i).Caption = ""
        frmMain.mnuQuickLink(i).Visible = False
    End If
  Next i
  frmMain.mnuQEdit.Enabled = Flag
  frmMain.mnuQSB.Visible = Flag
  If QuickListEditorOpen Then ShowQuickListEdit
End Sub

Sub ShowQuickListEdit()
Dim i As Integer, tabs As String
frmQuickList.List1.Clear
For i = 0 To 9
  If QuickList(i).Description <> "" Then
    If Len(QuickList(i).Description) > 25 Then
       tabs = Chr$(9)
    Else
       tabs = Chr$(9) & Chr$(9)
    End If
    frmQuickList.List1.AddItem QuickList(i).Description & ":" & tabs & GetFileFromPath(QuickList(i).File)
  End If
Next i
If frmQuickList.List1.ListCount >= 1 Then
   frmQuickList.List1.ListIndex = 0
   frmQuickList.cmdProperties.Enabled = True
   frmQuickList.cmdRemove.Enabled = True
   frmQuickList.cmdClear.Enabled = True
   frmQuickList.cmdOpen.Enabled = True
   frmQuickList.lblDesc = QuickList(0).Description
   frmQuickList.lblFile = QuickList(0).File
Else
   frmQuickList.cmdProperties.Enabled = False
   frmQuickList.cmdRemove.Enabled = False
   frmQuickList.cmdClear.Enabled = False
   frmQuickList.cmdOpen.Enabled = False
   frmQuickList.lblDesc = ""
   frmQuickList.lblFile = ""
End If
End Sub

Sub Sort(x As Integer, ni As Integer)
  
  ' Revised on 11/9/97

  ' Bubble Sort the Items
  
  Dim i As Integer, passNum  As Integer
  Dim temp1 As FolderType, temp2 As FolderType
  frmSort.Show
  DoEvents
  For passNum = 2 To ni - 1
      For i = 2 To ni - passNum
         Get #FreeNum, i, temp1
         Get #FreeNum, i + 1, temp2
         Select Case x
           Case 0
              temp1.Company = Decript(Left$(Trim$(temp1.Company), 1))
              temp2.Company = Decript(Left$(Trim$(temp2.Company), 1))
              If Trim$(temp1.Company) > Trim$(temp2.Company) Then Swap i, i + 1, FreeNum
           Case 1
              temp1.AName = Decript(Left$(Trim$(temp1.AName), 1))
              temp2.AName = Decript(Left$(Trim$(temp2.AName), 1))
              If Trim$(temp1.AName) > Trim$(temp2.AName) Then Swap i, i + 1, FreeNum
         End Select
      Next i
  Next passNum
  LoadList
  Unload frmSort
End Sub

Sub SortArray(Array1() As String, Array2() As Integer)

    Dim i As Long, j As Long, K As Long, l As Long, m As Long
    Dim N1 As Long, Z As Variant, temp As Variant
    Dim SortNum As Integer
    N1 = UBound(Array1): m = N1

Sort1:

    m = m \ 2
    If m = 0 Then
        Exit Sub
    End If

    j = 1

Sort2:

    K = N1 - m
    i = j

Sort3:

    l = i + m
    If LCase(Array(i)) <= LCase(Array1(l)) Then GoTo Sort4
    temp = Array1(i): Array1(i) = Array1(l): Array1(l) = temp
    temp = Array2(i): Array2(i) = Array2(l): Array2(l) = temp
    i = i - m
    If i < m Then GoTo Sort4
    GoTo Sort3

Sort4:

    j = j + 1
    If j > K Then GoTo Sort1
    GoTo Sort2

End Sub

Sub Swap(L1 As Variant, L2 As Variant, Freenum2 As Integer)
   ' Swap the locations of two records
   ' Last Revised on 11/9/97

   Dim temp As FolderType
   
   Get Freenum2, L2, Pd
   
   temp.Company = Pd.Company
   temp.AName = Pd.AName
   temp.Address = Pd.Address
   temp.City = Pd.City
   temp.State = Pd.State
   temp.Zip_Code = Pd.Zip_Code
   temp.Home_Phone = Pd.Home_Phone
   temp.Bus_Phone = Pd.Bus_Phone
   temp.Pager = Pd.Pager
   temp.Fax = Pd.Fax
   temp.E_Mail = Pd.E_Mail
   temp.WebPage = Pd.WebPage
   temp.Notes = Pd.Notes
   temp.Bookmark = Pd.Bookmark

       
   Get Freenum2, L1, Pd
   Put Freenum2, L2, Pd
   Put Freenum2, L1, temp
End Sub

Sub UpdatePrintBuffer()
  ' Update the print buffer window
  If Not PrintBufferOpen Then Exit Sub
  Screen.MousePointer = 11
  Dim i As Integer, Flag As Integer
  frmPrintBuffer.List1.Enabled = True
  frmPrintBuffer.List1.Clear
  For i = 0 To UBound(PrintBuffer)
      If PrintBuffer(i) <> 0 Then
        GetPD FreeNum, PrintBuffer(i), False
        frmPrintBuffer.List1.AddItem Decript(Trim$(Pd.Company)) & "  " & Decript(Trim$(Pd.AName))
        Flag = True
      End If
  Next i
  If i = 1 Then
     frmPrintBuffer.Label1 = "1 item in buffer"
  Else
     frmPrintBuffer.Label1 = Trim$(Str$(1)) & " items in buffer:"
  End If
  If Flag Then
     frmPrintBuffer.List1.ListIndex = 0
     frmPrintBuffer.cmdRemove.Enabled = True
     frmPrintBuffer.cmdSort.Enabled = True
     frmPrintBuffer.cmdView.Enabled = True
     frmPrintBuffer.cmdClear.Enabled = True
     frmPrintBuffer.cmdPrint.Enabled = True
  Else
     frmPrintBuffer.List1.AddItem "<No items in print buffer>"
     frmPrintBuffer.List1.ListIndex = 0
     frmPrintBuffer.List1.Enabled = False
     frmPrintBuffer.cmdRemove.Enabled = False
     frmPrintBuffer.cmdSort.Enabled = False
     frmPrintBuffer.cmdView.Enabled = False
     frmPrintBuffer.cmdClear.Enabled = False
     frmPrintBuffer.cmdPrint.Enabled = False
  End If
  DoEvents
  Screen.MousePointer = 0
End Sub


Public Sub MarkAllWindowsClosed()
    ViewWindowOpen = False
    SearchResultsOpen = False
    BookmarksOpen = False
    RecyclingBinOpen = False
    PrintBufferOpen = False
    QuickListEditorOpen = False
End Sub

' CloseAllWindows
' Purpose:  Closes all editing windows
' Created on: 2-11-1999
' Modified: 2-11-1999
Public Sub CloseAllWindows()
    If ViewWindowOpen Then Unload ViewWindow
    If SearchResultsOpen Then Unload frmLocateIt
    If BookmarksOpen Then Unload frmbookmarks
    If RecyclingBinOpen Then Unload frmRecycle
    If PrintBufferOpen Then Unload frmPrintBuffer
    If QuickListEditorOpen Then Unload frmQuickList
End Sub

Public Sub NotAvailiableInAlpha()
    MsgBox "This feature has not been implimented in this prerelease.", , "Alpha Release"
End Sub
