Attribute VB_Name = "WININI"

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ~              WININI Utility Code 32bit 1.0            ~
' ~             Copyright ©1997-1998 Rietta Software      ~
' ~                                                       ~
' ~     Release: 11/7/98                                  ~
' ~     E-Mail: riettasoftware@hotmail.com                ~
' ~     http://members.xoom.com/rietta                    ~
' ~                                                       ~
' ~     Tested with Visual Basic 4.0                      ~
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' MsgBox button declarations
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons
Global Const MB_ICONSTOP = 16          ' Critical frmssage
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning frmssage
Global Const MB_ICONINFORMATION = 64   ' Information frmssage

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' Mouse Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4
Global crlf  As String                 ' Caraige Return Line Feed

' Misc
Public Const OFN_OVERWRITEPROMPT = &H2&
Public Const OFN_HIDEREADONLY = &H4&


'**********************************
'**  API Function Declarations:

Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Public Declare Function GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function GetPrivateProfileSection Lib "Kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileStringSections& Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Function GetPrivateProfileInt& Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String)
Public Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)

' API Constants
Const LB_FINDSTRINGEXACT = &H1A2
Public Function RemoveCRLF(t As String) As String
    Dim temp As String, t2 As String, i As Integer
    temp = ""
    For i = 1 To Len(t)
        t2 = Mid$(t, i, 1)
        Select Case t2
            Case Chr$(13):
            Case Chr$(10):
            Case Else: temp = temp & t2
        End Select
    Next i
    RemoveCRLF = temp
End Function

Public Function VBFindInList(lst As Control, ByVal find As String) As Integer
    Dim i As Integer, sel As Integer
    sel = lst.ListIndex
    find = Trim$(LCase$(find))
    For i = 0 To lst.ListCount - 1
        lst.ListIndex = i
        If Trim$(LCase$(lst.text)) = find Then
            VBFindInList = i
            Exit Function
        End If
    Next i
    VBFindInList = -2
End Function

Function BracesToCRLF(s As String)
   ' Replace the { and } in s with carriage returns and linefeeds

   Dim i As Integer, Result As String, t As String
   Result = ""
   For i = 1 To Len(s)
       Select Case Mid$(s, i, 1)
           Case "{": t = Chr$(13)
           Case "}": t = Chr$(10)
           Case Else: t = Mid$(s, i, 1)
       End Select
       Result = Result & t
   Next i
   BracesToCRLF = Result

End Function

Function CapFirst(ByVal text As String) As String
    If Len(text) > 1 Then
        CapFirst = UCase$(Left$(text, 1)) & Mid$(text, 2)
    Else
        CapFirst = UCase$(text)
    End If
End Function

Sub CenterMe(frm As Form)
   ' Center the form passed as frm in the middle of the screen

   Dim X, Y                            ' New top, left for the form
 
   X = (Screen.Width - frm.Width) / 2
   Y = (Screen.Height - frm.Height) / 2
   frm.Move X, Y                       ' Change location of the form

End Sub

Function CriptString(Code As String, key As Integer) As String
   ' This function will encrypt or decrypt a string.
   ' If the encrypt key is 5 then the decrypt key is -5.
   ' (If you don't understand why this is, get an algebra book and learn
   ' algebraic addition.)

   Dim i As Integer, NumCode As Integer, Enc As String
   On Error Resume Next
   If Code <> "" Then
      'Scan String and evaluate encripted charictor
      Enc = ""
      For i = 1 To Len(Code)
         NumCode = Asc(Mid$(Code, i, 1))
         NumCode = NumCode + key    'Compute value for de/en--cripted charator
         Enc = Enc & Chr$(NumCode)  'Create an ASCII charactor
      Next i
      CriptString = Enc
   End If
End Function

Function CRLFToBraces(s As String) As String
  ' Replace the carriage returns and linefeeds in s with { and }

  Dim i As Integer, Result As String, t As String
  Result = ""
  For i = 1 To Len(s)
      Select Case Mid$(s, i, 1)
          Case Chr$(13): t = "{"
          Case Chr$(10): t = "}"
          Case Else: t = Mid$(s, i, 1)
      End Select
      Result = Result & t
  Next i
  CRLFToBraces = Result
End Function

Function GetFileFromPath(ByVal filepath As String) As String
  ' Get the file name from the specified path statement.
  ' for example, if filePath = "c:\myfiles\text\report.doc"
  ' then "report.doc" will be returned.

  Dim X As Integer, ct As Integer
  GetFileFromPath = filepath
  X = InStr(filepath, "\")
  Do While X
     ct = X
     X = InStr(ct + 1, filepath, "\")
  Loop
  If ct > 0 Then GetFileFromPath = Mid$(filepath, ct + 1)
End Function

Function GetMyIni(FileName As String, SectionName As String, keyname As String)
   ' Get information from a specified ini file.
   ' "filename" is the name of the ini file to read from.
   ' "appname" is the section from which you want to retrieve data.  A section
   ' name in and ini file is placed between two brackets (i.e. [Journal]).
   ' "keyname" is the keyword for the information under the specified section.
  
   Dim retVal, key, i, j
   Dim IniString As String

   ' This variable must be large enough to hold the return string
   ' from the GetPrivateProfileString API.
   IniString = String(255, 0)
   retVal = GetPrivateProfileString(SectionName, keyname, "", IniString, Len(IniString), FileName)
   GetMyIni = Left$(IniString, retVal)
End Function

Function GetPath(ByVal filepath As String)
  ' Get the path name from a path statement.
  ' for example, if filePath = "c:\myfiles\text\report.doc"
  ' then "c:\myfiles\text\" will be returned.
 
  Dim X As Integer, ct As Integer
  X = InStr(filepath, "\")
 
  Do While X
     ct = X
     X = InStr(ct + 1, filepath, "\")
  Loop
  If ct > 0 Then
      GetPath = Left$(filepath, ct)
  Else
      GetPath = ""   ' no path in string
  End If
End Function

Function GetWinIni(appname As String, keyname As String) As String
   ' Get information from win.ini file.
   ' "appname" is the section from which you want to retrieve data.  A section
   ' name in and ini file is placed between two brackets (i.e. [Journal]).
   ' "keyname" is the keyword for the information under the specified section.

   Dim ResultString As String * 128, temp As Long
   temp = GetProfileString(appname, keyname, "", ResultString, Len(ResultString))
   GetWinIni = Left$(ResultString, temp)
End Function

Public Function MyIni_GetInt(FileName As String, SectionName As String, keyname As String)
    Dim r As Long
    r = MyIni_GetLong(FileName, SectionName, keyname)
    MyIni_GetInt = Val(Format(r))
End Function

Sub SetCRLF()
   ' Set the CRLF variable to contain a carriage return and a linefeed

   crlf = Chr$(13) & Chr$(10)
End Sub

Sub SetMyIni(FileName As String, appname As String, keyname As String, stringtowrite As String)
   ' Save information to the specified ini file.
   ' "filename" is the filename of the ini file to which the data is to be saved.
   ' "appname" is the section from which you want to save data.  A section
   ' name in and ini file is placed between two brackets (i.e. [Journal]).
   ' "keyname" is the keyword for the information under the specified section.
   ' "KeyVal" is the information that is to be saved

   Dim temp As Long
   temp = WritePrivateProfileString(appname, keyname, stringtowrite, FileName)
End Sub

' Save information to the win.ini file.
' "appname" is the section from which you want to save data.  A section
' name in and ini file is placed between two brackets (i.e. [Journal]).
' "keyname" is the keyword for the information under the specified section.
' "KeyVal" is the information that is to be saved
'
Sub SetWinIni(appname As String, keyname As String, KeyVal As String)
   Dim temp As Long
   temp = WriteProfileString(appname, keyname, KeyVal)
End Sub


Sub VSwap(a As Variant, b As Variant)
  ' Visual Basic does not support the SWAP statement found in Quick Basic.
  ' This procedure simply simulates it.
  
  Dim temp As Variant
  temp = b
  b = a
  a = temp
End Sub


'
' Find if the specified section in an ini file
' has any keynames under it.  Returns a true or
' false value
'
Public Function MyIni_SectionExists(FileName As String, SectionName As String) As Integer

End Function

'
'  Loads the section names of an ini file into an array
'
Public Sub MyIni_LoadSections(FileName As String, iniSections() As String)
    Dim charNum As Long, sectionList As String
    Dim nullOffSet As Integer, ai As Integer
    
    sectionList = String$(128, 0)
    
    ' Windows 95/NT System Call
    charNum = GetPrivateProfileStringSections(0, 0, "", sectionList, 127, FileName)
    
    ReDim iniSections(0) As String
    ai = -1
    Do
        nullOffSet = InStr(sectionList, Chr$(0))
        If nullOffSet > 1 Then
            ai = ai + 1
            ReDim Preserve iniSections(ai) As String
            iniSections(ai) = Mid$(sectionList, 1, nullOffSet - 1)
            sectionList = Mid$(sectionList, nullOffSet + 1)
        End If
    Loop While nullOffSet > 1
End Sub

Public Function MyIni_GetLong(FileName As String, SectionName As String, keyname As String) As Integer
    Dim retVal As Long
    
    retVal = GetPrivateProfileInt(ByVal SectionName, ByVal keyname, 0, ByVal FileName)
    MyIni_GetLong = retVal
End Function

Public Function FindInList(lst As Control, searchText As String) As Long
    Dim hw&, t&
    hw& = lst.hwnd    ' Get the window handle for the list box
    t& = SendMessageByString&(hw&, LB_FINDSTRINGEXACT, -1, searchText)
    FindInList = t&
End Function
