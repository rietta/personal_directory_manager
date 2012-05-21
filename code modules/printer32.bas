Attribute VB_Name = "Print_Functions"
' ER Software Printer API
' This source was compiled from various refrences and examples

Option Explicit


Type typPrintInfo
    sName As String
    nSize As Long
    bBold As Boolean
    bItalic As Boolean
    bUnderline As Boolean
    bStrikethru As Boolean
    bPageNumbers As Boolean
    nColor As Double
    nColumns As Integer
End Type
Public HeadingProp As typPrintInfo, BodyProp As typPrintInfo


Const sWINDOWS_SECTION_NAME = "windows"
Const sDEVICES_SECTION_NAME = "devices"
Const sDEVICE_KEY_NAME = "device"

Type typeWindowsDevice
   sWindowsDeviceUserName As String
   sWindowsDeviceShortName As String
   sWindowsDevicePortName As String
End Type

Public DoPrint As Integer
Public SetAsDefault As Integer
Public PrintColor As Integer
Public PrintFont As String

'  Printer Dialog Flags

Public Const PD_ALLPAGES = &H0&
Public Const PD_SELECTION = &H1&
Public Const PD_PAGENUMS = &H2&
Public Const PD_NOSELECTION = &H4&
Public Const PD_NOPAGENUMS = &H8&
Public Const PD_COLLATE = &H10&
Public Const PD_PRINTTOFILE = &H20&
Public Const PD_PRINTSETUP = &H40&
Public Const PD_NOWARNING = &H80&
Public Const PD_RETURNDC = &H100&
Public Const PD_RETURNIC = &H200&
Public Const PD_RETURNDEFAULT = &H400&
Public Const PD_SHOWHELP = &H800&
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000


Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATPAINT = &HFB0A09
Public Const PATINVERT = &H5A0049
Public Const DSTINVERT = &H550009
Public Const BLACKNESS = &H42&
Public Const WHITENESS = &HFF0062
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&
Public Const DM_YRESOLUTION = &H2000&
Public Const DM_TTOPTION = &H4000&
Public Const DMORIENT_PORTRAIT = 1
Public Const DMORIENT_LANDSCAPE = 2
Public Const DMPAPER_LETTER = 1
Public Const DMPAPER_LETTERSMALL = 2
Public Const DMPAPER_TABLOID = 3
Public Const DMPAPER_LEDGER = 4
Public Const DMPAPER_LEGAL = 5
Public Const DMPAPER_STATEMENT = 6
Public Const DMPAPER_EXECUTIVE = 7
Public Const DMPAPER_A3 = 8
Public Const DMPAPER_A4 = 9
Public Const DMPAPER_A4SMALL = 10
Public Const DMPAPER_A5 = 11
Public Const DMPAPER_B4 = 12
Public Const DMPAPER_B5 = 13
Public Const DMPAPER_FOLIO = 14
Public Const DMPAPER_QUARTO = 15
Public Const DMPAPER_10X14 = 16
Public Const DMPAPER_11X17 = 17
Public Const DMPAPER_NOTE = 18
Public Const DMPAPER_ENV_9 = 19
Public Const DMPAPER_ENV_10 = 20
Public Const DMPAPER_ENV_11 = 21
Public Const DMPAPER_ENV_12 = 22
Public Const DMPAPER_ENV_14 = 23
Public Const DMPAPER_CSHEET = 24
Public Const DMPAPER_DSHEET = 25
Public Const DMPAPER_ESHEET = 26
Public Const DMPAPER_ENV_DL = 27
Public Const DMPAPER_ENV_C5 = 28
Public Const DMPAPER_ENV_C3 = 29
Public Const DMPAPER_ENV_C4 = 30
Public Const DMPAPER_ENV_C6 = 31
Public Const DMPAPER_ENV_C65 = 32
Public Const DMPAPER_ENV_B4 = 33
Public Const DMPAPER_ENV_B5 = 34
Public Const DMPAPER_ENV_B6 = 35
Public Const DMPAPER_ENV_ITALY = 36
Public Const DMPAPER_ENV_MONARCH = 37
Public Const DMPAPER_ENV_PERSONAL = 38
Public Const DMPAPER_FANFOLD_US = 39
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Public Const DMPAPER_USER = 256
Public Const DMBIN_UPPER = 1
Public Const DMBIN_ONLYONE = 1
Public Const DMBIN_LOWER = 2
Public Const DMBIN_MIDDLE = 3
Public Const DMBIN_MANUAL = 4
Public Const DMBIN_ENVELOPE = 5
Public Const DMBIN_ENVMANUAL = 6
Public Const DMBIN_AUTO = 7
Public Const DMBIN_TRACTOR = 8
Public Const DMBIN_SMALLFMT = 9
Public Const DMBIN_LARGEFMT = 10
Public Const DMBIN_LARGECAPACITY = 11
Public Const DMBIN_CASSETTE = 14
Public Const DMBIN_USER = 256
Public Const DMRES_DRAFT = -1
Public Const DMRES_LOW = -2
Public Const DMRES_MEDIUM = -3
Public Const DMRES_HIGH = -4
Public Const DMCOLOR_MONOCHROME = 1
Public Const DMCOLOR_COLOR = 2
Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_VERTICAL = 2
Public Const DMDUP_HORIZONTAL = 3
Public Const DMTT_BITMAP = 1
Public Const DMTT_DOWNLOAD = 2
Public Const DMTT_SUBDEV = 3
Public Const DM_UPDATE = 1
Public Const DM_COPY = 2
Public Const DM_PROMPT = 4
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = 8
Public Const DM_IN_PROMPT = 4
Public Const DM_OUT_BUFFER = 2
Public Const DM_OUT_DEFAULT = 1
Public Const DC_FIELDS = 1
Public Const DC_PAPERS = 2
Public Const DC_PAPERSIZE = 3
Public Const DC_MINEXTENT = 4
Public Const DC_MAXEXTENT = 5
Public Const DC_BINS = 6
Public Const DC_DUPLEX = 7
Public Const DC_SIZE = 8
Public Const DC_EXTRA = 9
Public Const DC_VERSION = 10
Public Const DC_DRIVER = 11
Public Const DC_BINNAMES = 12
Public Const DC_ENUMRESOLUTIONS = 13
Public Const DC_FILEDEPENDENCIES = 14
Public Const DC_TRUETYPE = 15
Public Const DC_PAPERNAMES = 16
Public Const DC_ORIENTATION = 17
Public Const DC_COPIES = 18
Public Const DCTT_BITMAP = &H1&
Public Const DCTT_DOWNLOAD = &H2&
Public Const DCTT_SUBDEV = &H4&
Public Const SP_NOTREPORTED = &H4000
Public Const SP_ERROR = (-1)
Public Const SP_APPABORT = (-2)
Public Const SP_USERABORT = (-3)
Public Const SP_OUTOFDISK = (-4)
Public Const SP_OUTOFMEMORY = (-5)
Public Const PR_JOBSTATUS = &H0
Public Const DRIVERVERSION = 0
Public Const TECHNOLOGY = 2
Public Const HORZSIZE = 4
Public Const VERTSIZE = 6
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const BITSPIXEL = 12
Public Const PLANES = 14
Public Const NUMBRUSHES = 16
Public Const NUMPENS = 18
Public Const NUMMARKERS = 20
Public Const NUMFONTS = 22
Public Const NUMCOLORS = 24
Public Const PDEVICESIZE = 26
Public Const CURVECAPS = 28
Public Const LINECAPS = 30
Public Const POLYGONALCAPS = 32
Public Const TEXTCAPS = 34
Public Const CLIPCAPS = 36
Public Const RASTERCAPS = 38
Public Const ASPECTX = 40
Public Const ASPECTY = 42
Public Const ASPECTXY = 44
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const SIZEPALETTE = 104
Public Const NUMRESERVED = 106
Public Const COLORRES = 108
Public Const RC_BITBLT = 1
Public Const RC_BANDING = 2
Public Const RC_SCALING = 4
Public Const RC_BITMAP64 = 8
Public Const RC_GDI20_OUTPUT = &H10
Public Const RC_DI_BITMAP = &H80
Public Const RC_PALETTE = &H100
Public Const RC_DIBTODEV = &H200
Public Const RC_BIGFONT = &H400
Public Const RC_STRETCHBLT = &H800
Public Const RC_FLOODFILL = &H1000
Public Const RC_STRETCHDIB = &H2000
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const DIB_RGB_COLORS = 0
Public Const DIB_PAL_COLORS = 1

' Public variables
Public AbortPrinting%
Public UseHourglass%

'  size of a device name string
Public Const CCHDEVICENAME = 32

'  size of a form name string
Public Const CCHFORMNAME = 32

Public Const BI_RGB = 0&

Type POINTAPI
        x As Long
        y As Long
End Type

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As Long
        DesiredAccess As Long
End Type

Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
End Type

Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

' BITMAPINFO for this example is for 16 color bitmap
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(15) As RGBQUAD
End Type

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public dmout As DEVMODE

Private Declare Function GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function OpenPrinterBynum Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, ByVal pDevModeInput As Long) As Long
Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hwnd As Long, ByVal FLAGS As Long) As Long
Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, ByVal lpDevMode As Long) As Long
Declare Function CreateDCBynum Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndDocAPI Lib "gdi32" Alias "EndDoc" (ByVal hdc As Long) As Long
Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SetAbortProc Lib "gdi32" (ByVal hdc As Long, ByVal lpAbortProc As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Sub EditPrinterProperties(hwnd As Long)
    Dim dm As DEVMODE
    Dim bufsize&, res&
    Dim dmInBuf() As Byte
    Dim dmOutBuf() As Byte
    Dim hPrinter&
    Dim DeviceName$
        
    hPrinter = OpenDefaultPrinter(DeviceName$)
    If hPrinter = 0 Then
        MsgBox "Unable to open default printer"
        Exit Sub
    End If

    ' The output DEVMODE structure will reflect any changes
    ' made by the printer setup dialog box.
    ' Note that no changes will be made to the default
    ' printer settings!
    bufsize = DocumentProperties(hwnd, hPrinter, DeviceName$, 0, 0, 0)
    ReDim dmInBuf(bufsize)
    ReDim dmOutBuf(bufsize)
    
    res = DocumentProperties(hwnd, hPrinter, DeviceName$, agGetAddressForObject(dmOutBuf(0)), agGetAddressForObject(dmInBuf(0)), DM_IN_PROMPT Or DM_OUT_BUFFER)
        
    ' Copy the data buffer into the DEVMODE structure
    agCopyData dmOutBuf(0), dmout, Len(dmout)
        
    ClosePrinter hPrinter
End Sub


'   This function retrieves the definition of the default
'   printer on this system
'
 Function GetDefPrinter$()
    Dim def$
    Dim di&

    def$ = String$(128, 0)
    di = GetProfileString("WINDOWS", "DEVICE", "", def$, 127)
    def$ = agGetStringFromLPSTR$(def$)
    GetDefPrinter$ = def$

End Function
'   This function returns the driver module name
'
Private Function GetDeviceDriver$(dev$)
    Dim firstpos%, nextpos%
    firstpos% = InStr(dev$, ",")
    nextpos% = InStr(firstpos% + 1, dev$, ",")
    GetDeviceDriver$ = Mid$(dev$, firstpos% + 1, nextpos% - firstpos% - 1)
End Function


'   Retrieves the name portion of a device string
'
 Function GetDeviceName$(dev$)
    Dim npos%
    npos% = InStr(dev$, ",")
    GetDeviceName$ = Left$(dev$, npos% - 1)
End Function


'   Returns the output destination for the specified device
'
 Function GetDeviceOutput$(dev$)
    Dim firstpos%, nextpos%
    firstpos% = InStr(dev$, ",")
    nextpos% = InStr(firstpos% + 1, dev$, ",")
    GetDeviceOutput$ = Mid$(dev$, nextpos% + 1)
End Function

 Function OpenDefaultPrinter(Optional DeviceName) As Long
    Dim dev$, devname$, devoutput$
    Dim hPrinter&, res&
    Dim pdefs As PRINTER_DEFAULTS
    
    pdefs.pDatatype = vbNullString
    pdefs.pDevMode = 0
    pdefs.DesiredAccess = PRINTER_ACCESS_ADMINISTER

    
    dev$ = GetDefPrinter$() ' Get default printer info
    
    If dev$ = "" Then Exit Function
    devname$ = GetDeviceName$(dev$)
    devoutput$ = GetDeviceOutput$(dev$)
    
    If Not IsMissing(DeviceName) Then
        DeviceName = devname$
    End If
    
    ' You can use OpenPrinterBynum to pass a zero as the
    ' third parameter, but you won't have full access to
    ' edit the printer properties
    res& = OpenPrinter(devname$, hPrinter, pdefs)
    If res <> 0 Then OpenDefaultPrinter = hPrinter
End Function


Sub GetDefaultPrinter(recDefaultPrinter As typeWindowsDevice)

'This routine returns the "default" Windows printer.

Dim nStrPos As Integer
Dim sDefaultPrinter As String
Dim nRC As Integer


sDefaultPrinter = sDUINI_GetString(sWINDOWS_SECTION_NAME, sDEVICE_KEY_NAME, "", "")
nStrPos = InStr(sDefaultPrinter, ",")

recDefaultPrinter.sWindowsDeviceUserName = Left$(sDefaultPrinter, nStrPos - 1)

sDefaultPrinter = Mid$(sDefaultPrinter, nStrPos + 1)
nStrPos = InStr(sDefaultPrinter, ",")

recDefaultPrinter.sWindowsDeviceShortName = Left$(sDefaultPrinter, nStrPos - 1)
recDefaultPrinter.sWindowsDevicePortName = Mid$(sDefaultPrinter, nStrPos + 1)

End Sub

Sub GetInstalledPrinters(recInstalledPrinters() As typeWindowsDevice)
    
'This routine enumerates the "installed" Windows printers.

Dim nStrPos As Integer, nPrtSub As Integer
Dim sInstalledPrinter As String
ReDim sPrinterNames(0) As String


Call DUINI_GetKeyNames(sDEVICES_SECTION_NAME, sPrinterNames(), "")

ReDim recInstalledPrinters(UBound(sPrinterNames))
    
For nPrtSub = 1 To UBound(sPrinterNames)
   sInstalledPrinter = sDUINI_GetString(sDEVICES_SECTION_NAME, sPrinterNames(nPrtSub), "", "")
   nStrPos = InStr(sInstalledPrinter, ",")

   recInstalledPrinters(nPrtSub).sWindowsDeviceUserName = sPrinterNames(nPrtSub)
   recInstalledPrinters(nPrtSub).sWindowsDeviceShortName = Left$(sInstalledPrinter, nStrPos - 1)

   sInstalledPrinter = Mid$(sInstalledPrinter, nStrPos + 1)
   nStrPos = InStr(sInstalledPrinter, ",")

   If nStrPos > 0 Then
      recInstalledPrinters(nPrtSub).sWindowsDevicePortName = Left$(sInstalledPrinter, nStrPos - 1)
   Else
      recInstalledPrinters(nPrtSub).sWindowsDevicePortName = sInstalledPrinter
   End If
Next

End Sub

Sub SetDefaultPrinter(recDefaultPrinter As typeWindowsDevice)
    
'This routine sets the "default" Windows printer to one
'of the "installed" printers.

Dim sNewPrinter As String
Dim nRC As Integer


sNewPrinter = recDefaultPrinter.sWindowsDeviceUserName + "," + recDefaultPrinter.sWindowsDeviceShortName + "," + recDefaultPrinter.sWindowsDevicePortName
    
Call DUINI_WriteString(sWINDOWS_SECTION_NAME, sDEVICE_KEY_NAME, sNewPrinter, "")
Call DUINI_BroadcastWININIChange(sWINDOWS_SECTION_NAME)

End Sub

'   Prints the bitmap in the picture1 control to the
'   printer context specified.
'
 Sub PrintBitmap(hdc&, Picture1 As PictureBox)
    Dim bi As BITMAPINFO
    Dim dctemp&, dctemp2&
    Dim msg$
    Dim bufsize&
    Dim bm As BITMAP
    Dim ghnd&
    Dim gptr&
    Dim xpix&, ypix&
    Dim doscale&
    Dim uy&, ux&
    Dim di&

    ' Create a temporary memory DC and select into it
    ' the background picture of the picture1 control.
    dctemp& = CreateCompatibleDC(Picture1.hdc)
    
    ' Get the size of the picture bitmap
    di = GetObjectAPI(Picture1.Picture, Len(bm), bm)

    ' Can this printer handle the DIB?
    If (GetDeviceCaps(dctemp, RASTERCAPS)) And RC_DIBTODEV = 0 Then
        msg$ = "This device does not support DIB's" + vbCrLf + "See source code for further info"
        MsgBox msg$, 0, "No DIB support"
    End If

    ' Fill the BITMAPINFO for the desired DIB
    bi.bmiHeader.biSize = Len(bi.bmiHeader)
    bi.bmiHeader.biWidth = bm.bmWidth
    bi.bmiHeader.biHeight = bm.bmHeight
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 4
    bi.bmiHeader.biCompression = BI_RGB
    ' Now calculate the data buffer size needed
    bufsize& = bi.bmiHeader.biWidth

    ' Figure out the number of bytes based on the
    ' number of pixels in each byte. In this case we
    ' really don't need all this code because this example
    ' always uses a 16 color DIB, but the code is shown
    ' here for your future reference
    Select Case bi.bmiHeader.biBitCount
        Case 1
            bufsize& = (bufsize& + 7) / 8
        Case 4
            bufsize& = (bufsize& + 1) / 2
        Case 24
            bufsize& = bufsize& * 3
    End Select
    ' And make sure it aligns on a long boundary
    bufsize& = ((bufsize& + 3) / 4) * 4
    ' And multiply by the # of scan lines
    bufsize& = bufsize& * bi.bmiHeader.biHeight

    ' Now allocate a buffer to hold the data
    ' We use the global memory pool because this buffer
    ' could easily be above 64k bytes.
    ghnd = GlobalAlloc(GMEM_MOVEABLE, bufsize&)
    gptr& = GlobalLock&(ghnd)

    di = GetDIBits(dctemp, Picture1.Picture, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)
    di = SetDIBitsToDevice(hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)

    ' Now see if we can also print a scaled version
    xpix = GetDeviceCaps(hdc, HORZRES)
    ' We subtract off the size of the bitmap already
    ' printed, plus some extra space
    ypix = GetDeviceCaps(hdc, VERTRES) - (bm.bmHeight + 50)

    ' Find out the largest multiplier we can use and still
    ' fit on the page
    doscale = xpix / bm.bmWidth
    If (ypix / bm.bmHeight < doscale) Then doscale = ypix / bm.bmHeight
    If doscale > 1 Then
        ux = bm.bmWidth * doscale
        uy = bm.bmHeight * doscale
        ' Now how this is offset a bit so that we don't
        ' print over the 1:1 scaled bitmap
        di = StretchDIBits(hdc, 0, bm.bmHeight + 50, ux, uy, 0, 0, bm.bmWidth, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS, SRCCOPY)
    End If
    ' Dump the global memory block
    di = GlobalUnlock(ghnd)
    di = GlobalFree(ghnd)
    di = DeleteDC(dctemp)

End Sub

' Shows information about the current device mode
'
Sub ShowDevMode(dm As DEVMODE)
    Dim crlf$
    Dim a$

    crlf$ = Chr$(13) + Chr$(10)
    a$ = "Device name = " + agGetStringFromLPSTR$(dm.dmDeviceName) + crlf$
    a$ = a$ + "Devmode Version: " + Hex$(dm.dmSpecVersion) + ", Driver version: " + Hex$(dm.dmDriverVersion) + crlf$
    a$ = a$ + "Orientation: "
    If dm.dmOrientation = DMORIENT_PORTRAIT Then a$ = a$ + "Portrait" Else a$ = a$ + "Landscape"
    a$ = a$ + crlf$
    a$ = a$ + "Field mask = " + Hex$(dm.dmFields) + crlf$
    a$ = a$ + "Copies = " + Str$(dm.dmCopies) + crlf$
    If dm.dmFields And DM_YRESOLUTION <> 0 Then
        a$ = a$ + "X,Y resolution = " + Str$(dm.dmPrintQuality) + "," + Str$(dm.dmYResolution) + crlf$
    End If
    MsgBox a$, 0, "Devmode structure"
End Sub


Function GetCurrentPrinter() As String
 Dim DefPrinter As typeWindowsDevice
 GetDefaultPrinter DefPrinter
 GetCurrentPrinter = DefPrinter.sWindowsDeviceUserName & " on " & DefPrinter.sWindowsDevicePortName
End Function

Function PrinterSetup(cmdlg As Control) As Integer
 cmdlg.Copies = 1
 cmdlg.PrinterDefault = True
 cmdlg.FLAGS = PD_PRINTSETUP
 cmdlg.Action = 5
End Function


Sub SavePrintFormatting()

    SaveSetting "PDM", "Binder Print Properties", "Heading Font", HeadingProp.sName
    SaveSetting "PDM", "Binder Print Properties", "Heading Size", HeadingProp.nSize
    SaveSetting "PDM", "Binder Print Properties", "Heading Bold", HeadingProp.bBold
    SaveSetting "PDM", "Binder Print Properties", "Heading Italic", HeadingProp.bItalic
    SaveSetting "PDM", "Binder Print Properties", "Heading Underline", HeadingProp.bUnderline
    SaveSetting "PDM", "Binder Print Properties", "Heading Strikethru", HeadingProp.bStrikethru
    SaveSetting "PDM", "Binder Print Properties", "Heading Color", HeadingProp.nColor
    SaveSetting "PDM", "Binder Print Properties", "Page Numbers", HeadingProp.bPageNumbers
    
    SaveSetting "PDM", "Binder Print Properties", "Body Font", BodyProp.sName
    SaveSetting "PDM", "Binder Print Properties", "Body Size", BodyProp.nSize
    SaveSetting "PDM", "Binder Print Properties", "Body Bold", BodyProp.bBold
    SaveSetting "PDM", "Binder Print Properties", "Body Italic", BodyProp.bItalic
    SaveSetting "PDM", "Binder Print Properties", "Body Underline", BodyProp.bUnderline
    SaveSetting "PDM", "Binder Print Properties", "Body Strikethru", BodyProp.bStrikethru
    SaveSetting "PDM", "Binder Print Properties", "Body Color", BodyProp.nColor
    SaveSetting "PDM", "Binder Print Properties", "Columns", BodyProp.nColumns
    

End Sub


Sub LoadPrintFormatting()
    
    DefaultPrintSettings
    
    On Error Resume Next
    HeadingProp.sName = GetSetting("PDM", "Binder Print Properties", "Heading Font")
    HeadingProp.nSize = GetSetting("PDM", "Binder Print Properties", "Heading Size")
    HeadingProp.bBold = GetSetting("PDM", "Binder Print Properties", "Heading Bold")
    HeadingProp.bItalic = GetSetting("PDM", "Binder Print Properties", "Heading Italic")
    HeadingProp.bUnderline = GetSetting("PDM", "Binder Print Properties", "Heading Underline")
    HeadingProp.bStrikethru = GetSetting("PDM", "Binder Print Properties", "Heading Strikethru")
    HeadingProp.nColor = GetSetting("PDM", "Binder Print Properties", "Heading Color")
    HeadingProp.bPageNumbers = GetSetting("PDM", "Binder Print Properties", "Page Numbers")
    
    BodyProp.sName = GetSetting("PDM", "Binder Print Properties", "Body Font")
    BodyProp.nSize = GetSetting("PDM", "Binder Print Properties", "Body Size")
    BodyProp.bBold = GetSetting("PDM", "Binder Print Properties", "Body Bold")
    BodyProp.bItalic = GetSetting("PDM", "Binder Print Properties", "Body Italic")
    BodyProp.bUnderline = GetSetting("PDM", "Binder Print Properties", "Body Underline")
    BodyProp.bStrikethru = GetSetting("PDM", "Binder Print Properties", "Body Strikethru")
    BodyProp.nColor = GetSetting("PDM", "Binder Print Properties", "Body Color")
    BodyProp.nColumns = GetSetting("PDM", "Binder Print Properties", "Columns")
  

End Sub

Sub DefaultPrintSettings()

    HeadingProp.sName = "Arial"
    HeadingProp.nSize = 12
    HeadingProp.bBold = True
    HeadingProp.bItalic = False
    HeadingProp.bUnderline = True
    HeadingProp.bStrikethru = False
    HeadingProp.nColor = vbBlack
    HeadingProp.bPageNumbers = True
    
    BodyProp.sName = "Times New Roman"
    BodyProp.nSize = 12
    BodyProp.bBold = False
    BodyProp.bItalic = False
    BodyProp.bUnderline = False
    BodyProp.bStrikethru = False
    BodyProp.nColor = vbBlack
    BodyProp.bPageNumbers = False
    BodyProp.nColumns = 1
    
End Sub



