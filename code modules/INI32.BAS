Attribute VB_Name = "Ini_Functions"
Option Explicit

Private Declare Function nWinAPI_GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpDirName As String, ByVal nSize As Integer) As Long

Private Declare Function nWinAPI_WriteProfileString Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpstring As String) As Long
Private Declare Function nWinAPI_WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpstring As String, ByVal lpFileName As String) As Long

Private Declare Function nWinAPI_GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Long
Private Declare Function nWinAPI_GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
Private Declare Function nWinAPI_GetProfileKeys Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Long, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Long
Private Declare Function nWinAPI_GetPrivateProfileKeys Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Long, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long

Private Declare Function nWinAPI_DeleteProfileKey Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpstring As Long) As Long
Private Declare Function nWinAPI_DeletePrivateProfileKey Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpstring As Long, ByVal lpFileName As String) As Long
Private Declare Function nWinAPI_DeleteProfileSection Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Long, ByVal lpstring As Long) As Long
Private Declare Function nWinAPI_DeletePrivateProfileSection Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Long, ByVal lpstring As Long, ByVal lpFileName As String) As Long

Private Declare Function nWinAPI_BroadcastWININIChange Lib "User32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Long

Const nHWND_BROADCAST = &HFFFF
Const nWM_WININICHANGE = &H1A

Sub DUINI_BroadcastWININIChange(sSectionName As String)

'This routine notifies all running tasks of a change
'to a WIN.INI setting.

Dim nRC As Long


nRC = nWinAPI_BroadcastWININIChange(nHWND_BROADCAST, nWM_WININICHANGE, 0, sSectionName)

End Sub

Sub DUINI_DeleteKey(sSectionName As String, sKeyName As String, sProfileName As String)

'This routine removes a keyname, and its associated value.

Dim nRC As Long


If Trim$(sProfileName) = "" Then
   nRC = nWinAPI_DeleteProfileKey(sSectionName, sKeyName, 0)
Else
   nRC = nWinAPI_DeletePrivateProfileKey(sSectionName, sKeyName, 0, sProfileName)
End If

End Sub

Sub DUINI_DeleteSection(sSectionName As String, sProfileName As String)
'This routine removes an entire section.  All keynames
'and their associated values are deleted.

Dim nRC As Long


If Trim$(sProfileName) = "" Then
   nRC = nWinAPI_DeleteProfileSection(sSectionName, 0, 0)
Else
   nRC = nWinAPI_DeletePrivateProfileSection(sSectionName, 0, 0, sProfileName)
End If

End Sub

Sub DUINI_GetKeyNames(sSectionName As String, sKeyNames() As String, sProfileName As String)

'This routine enumerates all the keynames present in a
'given section.

Dim nStrPos As Integer, nKeyCount As Integer, nStart As Integer
Dim nKeyNamesLength As Long
Dim sKeyNameString As String


sKeyNameString = Space$(1024)

If Trim$(sProfileName) = "" Then
   nKeyNamesLength = nWinAPI_GetProfileKeys(sSectionName, 0, "", sKeyNameString, Len(sKeyNameString))
Else
   nKeyNamesLength = nWinAPI_GetPrivateProfileKeys(sSectionName, 0, "", sKeyNameString, Len(sKeyNameString), sProfileName)
End If

nKeyCount = 0
ReDim sKeyNames(0)

If nKeyNamesLength > 0 Then
   sKeyNameString = Left$(sKeyNameString, nKeyNamesLength)

   If right$(sKeyNameString, 1) <> Chr$(0) Then
      sKeyNameString = sKeyNameString + Chr$(0)
   End If

   nKeyNamesLength = Len(sKeyNameString)
   nStart = 1
   Do
      nStrPos = InStr(nStart, sKeyNameString, Chr$(0))

      If nStrPos > 0 Then
         nKeyCount = nKeyCount + 1
         ReDim Preserve sKeyNames(nKeyCount)
         sKeyNames(nKeyCount) = Mid$(sKeyNameString, nStart, nStrPos - nStart)

         If nStrPos < nKeyNamesLength Then
            nStart = nStrPos + 1
         Else
            Exit Do
         End If
      Else
         Exit Do
      End If
   Loop
End If

End Sub


Sub DUINI_WriteString(sSectionName As String, sKeyName As String, sKeyValue As String, sProfileName As String)

'This routine updates the value associated with a given
'keyname, within a given section.  If either the keyname,
'the section, or the .INI file itself is not present,
'they are created.

Dim nRC As Integer


If Trim$(sProfileName) = "" Then
   nRC = nWinAPI_WriteProfileString(sSectionName, sKeyName, sKeyValue)
Else
   nRC = nWinAPI_WritePrivateProfileString(sSectionName, sKeyName, sKeyValue, sProfileName)
End If

End Sub

Function sDUINI_GetString(sSectionName As String, sKeyName As String, sDefaultValue As String, sProfileName As String) As String

'This function return the value associated with a given
'keyname, within a given section.  If either the keyname,
'the section, or the .INI file itself is not present,
'the "default" value is returned instead.

Dim nKeyValueLength As Integer
Dim sKeyValue As String


sKeyValue = Space$(256)

If Trim$(sProfileName) = "" Then
   nKeyValueLength = nWinAPI_GetProfileString(sSectionName, sKeyName, sDefaultValue, sKeyValue, Len(sKeyValue))
Else
   nKeyValueLength = nWinAPI_GetPrivateProfileString(sSectionName, sKeyName, sDefaultValue, sKeyValue, Len(sKeyValue), sProfileName)
End If

sDUINI_GetString = Left$(sKeyValue, nKeyValueLength)

End Function

