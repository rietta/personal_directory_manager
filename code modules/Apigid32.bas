Attribute VB_Name = "DesawareAPIGuide32"
Option Explicit
' ------------------------------------------------------------------------
'
'     APIGID32.BAS -- APIGID32.DLL API Declarations for Visual Basic
'
'                       Copyright (C) 1992-1996 Desaware
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Desaware and Ziff-Davis Press has no
'  warranty, obligation or liability for its contents.
'  Refer to the Ziff-Davis Visual Basic Programmer's Guide to the
'  Win32 API for further information.
'
' ------------------------------------------------------------------------
Type POINTS
        x  As Integer
        y  As Integer
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

#If Win32 Then
Declare Function agGetInstance& Lib "apigid32.dll" ()
Declare Function agPOINTStoLong& Lib "apigid32.dll" (pt As POINTS)
Declare Sub agCopyData Lib "apigid32.dll" (source As Any, dest As Any, ByVal nCount&)
Declare Sub agCopyDataBynum Lib "apigid32.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function agGetAddressForObject& Lib "apigid32.dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "apigid32.dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "apigid32.dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "apigid32.dll" Alias "agGetAddressForObject" (ByVal lpstring$) ' See warning!
Declare Function agGetAddressForVBString& Lib "apigid32.dll" (vbstring$)
Declare Function agGetStringFrom2NullBuffer$ Lib "apigid32.dll" (ByVal ptr&)
Declare Function agGetStringFromLPSTR$ Lib "apigid32.dll" (ByVal src$)
Declare Function agGetStringFromPointer$ Lib "apigid32.dll" Alias "agGetStringFromLPSTR" (ByVal ptr&)
Declare Function agSwapBytes% Lib "apigid32.dll" (ByVal src%)
Declare Function agSwapWords& Lib "apigid32.dll" (ByVal src&)
Declare Function agMakeROP4& Lib "apigid32.dll" (ByVal foreground&, ByVal background&)
Declare Function agGetWndInstance& Lib "apigid32.dll" (ByVal hwnd&)
Declare Function agDWORDto2Integers& Lib "apigid32.dll" (ByVal l&, lw%, lh%)
Declare Function agIsValidName& Lib "apigid32.dll" (ByVal o As Object, ByVal lpname$)
Declare Function agInp% Lib "apigid32.dll" (ByVal portid%)
Declare Function agInpw% Lib "apigid32.dll" (ByVal portid%)
Declare Function agInpd& Lib "apigid32.dll" (ByVal portid%)
Declare Sub agOutp Lib "apigid32.dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "apigid32.dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpd Lib "apigid32.dll" (ByVal portid%, ByVal outval&)

' Declared As Any to allow it to be used within classes, not to mention by other
' double long structures
Declare Sub agSubtractFileTimes Lib "apigid32.dll" (f1 As Any, f2 As Any, f3 As Any)
#Else
' Note, not all 16 bit declarations have equivalent 32 bit functions
' and vice versa. Nor is their behavior always identical.
' Refer to the Visual Basic Programmer's Guide to the Windows API (16 bit)
' for documentation on the following functions

Global Const CTLFLG_USESPALETTE% = 2
Global Const CTLFLG_HASPALETTE% = 1
 

Declare Function agGetControlHwnd% Lib "Apiguide.dll" (hctl As Control)
Declare Function agGetInstance% Lib "Apiguide.dll" ()
Declare Sub agCopyData Lib "Apiguide.dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "Apiguide.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Function agGetAddressForObject& Lib "Apiguide.dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "Apiguide.dll" Alias "agGetAddressForObject" (ByVal lpstring$)
Declare Function agGetAddressForVBString& Lib "Apiguide.dll" (vbstring$)
Declare Function agGetStringFromLPSTR$ Lib "Apiguide.dll" (ByVal lpstring$)
Declare Function agGetControlName$ Lib "Apiguide.dll" (ByVal hwnd%)
Declare Function agPOINTAPItoLong& Lib "Apiguide.dll" (pt As POINTAPI)
Declare Function agPOINTStoLong& Lib "Apiguide.dll" Alias "agPOINTAPItoLong" (pt As POINTS)
Declare Sub agDWORDto2Integers Lib "Apiguide.dll" (ByVal l&, lw%, lh%)
Declare Function agXPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
Declare Function agYTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
Declare Function agDeviceCapabilities& Lib "Apiguide.dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "Apiguide.dll" (ByVal hwnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "Apiguide.dll" (ByVal hwnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "Apiguide.dll" (ByVal portid%)
Declare Function agInpw% Lib "Apiguide.dll" (ByVal portid%)
Declare Sub agOutp Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
Declare Function agHugeOffset& Lib "Apiguide.dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "Apiguide.dll" ()
Declare Function agVBSendControlMsg& Lib "Apiguide.dll" (ctl As Control, ByVal msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "Apiguide.dll" (ctl As Control, ByVal mask&, ByVal value&)
Declare Sub agVBScreenToClient Lib "Apiguide.dll" (ctl As Control, pap As POINTS)
Declare Sub agVBClientToScreen Lib "Apiguide.dll" (ctl As Control, pap As POINTS)
Declare Function dwVBSetControlFlags& Lib "Apiguide.dll" (ctl As Control, ByVal mask&, ByVal value&)

#End If

