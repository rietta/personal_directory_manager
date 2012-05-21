Attribute VB_Name = "ThreeD_Effects_Generator"
Option Explicit


'**********************************
'**  Function Declarations:

#If Win32 Then
Private Declare Function GetActiveWindow& Lib "User32" ()
Private Declare Function GetWindowDC& Lib "User32" (ByVal hwnd As Long)
Private Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Private Declare Function ReleaseDC& Lib "User32" (ByVal hwnd As Long, ByVal hdc As Long)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long)
'Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI)
#End If 'WIN32



'Pen Styles
Global Const PS_SOLID = 0
Global Const PS_DASH = 1        '  -------
Global Const PS_DOT = 2         '  .......
Global Const PS_DASHDOT = 3     '  _._._._
Global Const PS_DASHDOTDOT = 4  '  _.._.._
Global Const PS_NULL = 5
Global Const PS_INSIDEFRAME = 6

'Colors
Global Const COLOR_BLACK = &H0&
Global Const COLOR_LIGHT_GRAY = &HC0C0C0
Global Const COLOR_DARK_GRAY = &H808080
Global Const COLOR_WHITE = &HFFFFFF

'Misc Consts
Global Const dlgBorderWidth = 6
Global Const dlgCaptionHeight = 18
Global Const dlgSysMenuWidth = 18



' Types, Declares and Const needed for Dlg3DBlaster
Type typRect
    Left As Integer
    Top As Integer
    right As Integer
    bottom As Integer
End Type

Declare Function SetTextAlign Lib "GDI" (ByVal hdc As Integer, ByVal wFlags As Integer) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "GDI" (ByVal hdc As Integer, ByVal nBkMode As Integer) As Integer
Declare Function GetStockObject Lib "GDI" (ByVal nIndex As Integer) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal lpstring As String, ByVal nCount As Integer) As Integer

Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function FillRect Lib "User" (ByVal hdc As Integer, lpRect As typRect, ByVal hBrush As Integer) As Integer
Declare Function Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer

Declare Function GetSysColor Lib "User" (ByVal nIndex As Integer) As Long

Global Const WM_NCACTIVATE = &H86
Global Const WM_NCPAINT = &H85

Global Const TA_CENTER = 6
Global Const TA_TOP = 0

Global Const COLOR_ACTIVECAPTION = 2
Global Const COLOR_CAPTIONTEXT = 9
Global Const COLOR_INACTIVECAPTION = 3
Global Const COLOR_INACTIVECAPTIONTEXT = 19

'Global Const TRANSPARENT = 1

Global Const SYSTEM_FONT = 13

'
' ComboBoxIn3D paints a 3D-border around the ComboBox given in ctrlCombo.
' frmTarget is the Form containing that ComboBox.
' If it is a Dropdown ComboBox, the ugly white space between the text field
' and the Dropdown-button is painted in lightgray.
'
'
Sub ComboBoxIn3D(ctrlCombo As Control, nBevel As Integer)
    Dim PixelX As Integer, PixelY As Integer
    Dim CTop As Integer, CRight As Integer, CBottom As Integer

    ' Just put "No 3D" in the Tag property and your ComboBox keeps 2D
    If InStr(UCase(ctrlCombo.Tag), "NO 3D") = 0 Then

        ControlIn3D ctrlCombo, nBevel, 0, True

        If ctrlCombo.Style = 0 Then             'Remove white space only
            PixelX = screen.TwipsPerPixelX      'if it is a Dropdown ComboBox
            PixelY = screen.TwipsPerPixelY
            CTop = ctrlCombo.Top
            CRight = ctrlCombo.Left + ctrlCombo.Width
            CBottom = ctrlCombo.Top + ctrlCombo.Height
            ctrlCombo.Parent.Line (CRight - PixelX * 24, CTop)-(CRight - PixelX * 18, CBottom - PixelY), COLOR_LIGHT_GRAY, BF
        End If
    End If

End Sub

'
' ControlIn3D paints a 3D-border around the control given in ctrlTarget.
' nBevel controls the the deepness, nSpace the distance between the control
' and the 3D-border and bInset sets the border to be drawn inset or outset.
'
' Parts of this code are taken from the VB Tips & Tricks help file.
' Original code written by Matej Nastran.
'
'
Sub ControlIn3D(ctrlTarget As Control, nBevel As Integer, nSpace As Integer, bInset As Integer)
    Dim CTop As Integer, CLeft As Integer, CRight As Integer, CBottom As Integer
    Dim PixelX As Integer, PixelY As Integer, AddX As Integer, AddY As Integer
    Dim i As Integer

    ' Just put "No 3D" in the Tag property and your control keeps 2D
    If InStr(UCase(ctrlTarget.Tag), "NO 3D") = 0 Then
        PixelX = screen.TwipsPerPixelX
        PixelY = screen.TwipsPerPixelY
        CTop = ctrlTarget.Top - PixelY
        CLeft = ctrlTarget.Left - PixelX
        CRight = ctrlTarget.Left + ctrlTarget.Width
        CBottom = ctrlTarget.Top + ctrlTarget.Height
        If bInset Then          ' Draw border inset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                ctrlTarget.Parent.Line (CLeft - AddX, CTop - AddY)-(CRight + AddX, CTop - AddY), COLOR_DARK_GRAY
                ctrlTarget.Parent.Line (CLeft - AddX, CTop - AddY)-(CLeft - AddX, CBottom + AddY), COLOR_DARK_GRAY
                ctrlTarget.Parent.Line (CLeft - AddX, CBottom + AddY)-(CRight + AddX + PixelX, CBottom + AddY), COLOR_WHITE
                ctrlTarget.Parent.Line (CRight + AddX, CTop - AddY)-(CRight + AddX, CBottom + AddY), COLOR_WHITE
            Next i
        Else                    ' Draw border outset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                ctrlTarget.Parent.Line (CRight + AddX, CBottom + AddY)-(CRight + AddX, CTop - AddY), COLOR_DARK_GRAY
                ctrlTarget.Parent.Line (CRight + AddX, CBottom + AddY)-(CLeft - AddX, CBottom + AddY), COLOR_DARK_GRAY
                ctrlTarget.Parent.Line (CRight + AddX, CTop - AddY)-(CLeft - AddX - PixelX, CTop - AddY), COLOR_WHITE
                ctrlTarget.Parent.Line (CLeft - AddX, CBottom + AddY)-(CLeft - AddX, CTop - AddY), COLOR_WHITE
            Next i
        End If
    End If

End Sub

Sub Controlin3DonPicbox(ctrlTarget As Control, nBevel As Integer, nSpace As Integer, bInset As Integer, picbox As PictureBox)
    Dim CTop As Integer, CLeft As Integer, CRight As Integer, CBottom As Integer
    Dim PixelX As Integer, PixelY As Integer, AddX As Integer, AddY As Integer
    Dim i As Integer

    ' Just put "No 3D" in the Tag property and your control keeps 2D
    If InStr(UCase(ctrlTarget.Tag), "NO 3D") = 0 Then
        PixelX = screen.TwipsPerPixelX
        PixelY = screen.TwipsPerPixelY
        CTop = ctrlTarget.Top - PixelY
        CLeft = ctrlTarget.Left - PixelX
        CRight = ctrlTarget.Left + ctrlTarget.Width
        CBottom = ctrlTarget.Top + ctrlTarget.Height
        If bInset Then          ' Draw border inset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                picbox.Line (CLeft - AddX, CTop - AddY)-(CRight + AddX, CTop - AddY), COLOR_DARK_GRAY
                picbox.Line (CLeft - AddX, CTop - AddY)-(CLeft - AddX, CBottom + AddY), COLOR_DARK_GRAY
                picbox.Line (CLeft - AddX, CBottom + AddY)-(CRight + AddX + PixelX, CBottom + AddY), COLOR_WHITE
                picbox.Line (CRight + AddX, CTop - AddY)-(CRight + AddX, CBottom + AddY), COLOR_WHITE
            Next i
        Else                    ' Draw border outset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                picbox.Line (CRight + AddX, CBottom + AddY)-(CRight + AddX, CTop - AddY), COLOR_DARK_GRAY
                picbox.Line (CRight + AddX, CBottom + AddY)-(CLeft - AddX, CBottom + AddY), COLOR_DARK_GRAY
                picbox.Line (CRight + AddX, CTop - AddY)-(CLeft - AddX - PixelX, CTop - AddY), COLOR_WHITE
                picbox.Line (CLeft - AddX, CBottom + AddY)-(CLeft - AddX, CTop - AddY), COLOR_WHITE
            Next i
        End If
    End If
End Sub



Public Sub Erase3DBorderOnPicBox(ctrlTarget As Control, nBevel As Integer, nSpace As Integer, bInset As Integer, picbox As PictureBox, color As Long)
    Dim CTop As Integer, CLeft As Integer, CRight As Integer, CBottom As Integer
    Dim PixelX As Integer, PixelY As Integer, AddX As Integer, AddY As Integer
    Dim i As Integer

    ' Just put "No 3D" in the Tag property and your control keeps 2D
    If InStr(UCase(ctrlTarget.Tag), "NO 3D") = 0 Then
        PixelX = screen.TwipsPerPixelX
        PixelY = screen.TwipsPerPixelY
        CTop = ctrlTarget.Top - PixelY
        CLeft = ctrlTarget.Left - PixelX
        CRight = ctrlTarget.Left + ctrlTarget.Width
        CBottom = ctrlTarget.Top + ctrlTarget.Height
        If bInset Then          ' Draw border inset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                picbox.Line (CLeft - AddX, CTop - AddY)-(CRight + AddX, CTop - AddY), color
                picbox.Line (CLeft - AddX, CTop - AddY)-(CLeft - AddX, CBottom + AddY), color
                picbox.Line (CLeft - AddX, CBottom + AddY)-(CRight + AddX + PixelX, CBottom + AddY), color
                picbox.Line (CRight + AddX, CTop - AddY)-(CRight + AddX, CBottom + AddY), color
            Next i
        Else                    ' Draw border outset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                picbox.Line (CRight + AddX, CBottom + AddY)-(CRight + AddX, CTop - AddY), color
                picbox.Line (CRight + AddX, CBottom + AddY)-(CLeft - AddX, CBottom + AddY), color
                picbox.Line (CRight + AddX, CTop - AddY)-(CLeft - AddX - PixelX, CTop - AddY), color
                picbox.Line (CLeft - AddX, CBottom + AddY)-(CLeft - AddX, CTop - AddY), color
            Next i
        End If
    End If

End Sub


'
' LineIn3D paints the given Line-control ctrlLine in 3D.
' frmTarget is the Form containing that Line.
'
Sub LineIn3D(ctrlLine As Control)

    If InStr(UCase(ctrlLine.Tag), "NO 3D") = 0 Then
        ctrlLine.BorderColor = COLOR_DARK_GRAY
        'Check if line is vertical or horizontal
        If Abs(ctrlLine.X2 - ctrlLine.X1) > Abs(ctrlLine.Y2 - ctrlLine.Y1) Then
            ctrlLine.Parent.Line (ctrlLine.X1, ctrlLine.Y1 + screen.TwipsPerPixelY)-(ctrlLine.X2, ctrlLine.Y2 + screen.TwipsPerPixelY), COLOR_WHITE
        Else
            ctrlLine.Parent.Line (ctrlLine.X1 + screen.TwipsPerPixelX, ctrlLine.Y1)-(ctrlLine.X2 + screen.TwipsPerPixelX, ctrlLine.Y2), COLOR_WHITE
        End If
    End If

End Sub

Sub LineIn3DonPicBox(picbox As PictureBox, ctrlLine As Control)
    If InStr(UCase(ctrlLine.Tag), "NO 3D") = 0 Then
        ctrlLine.BorderColor = COLOR_DARK_GRAY
        'Check if line is vertical or horizontal
        If Abs(ctrlLine.X2 - ctrlLine.X1) > Abs(ctrlLine.Y2 - ctrlLine.Y1) Then
            picbox.Line (ctrlLine.X1, ctrlLine.Y1 + screen.TwipsPerPixelY)-(ctrlLine.X2, ctrlLine.Y2 + screen.TwipsPerPixelY), COLOR_WHITE
        Else
            picbox.Line (ctrlLine.X1 + screen.TwipsPerPixelX, ctrlLine.Y1)-(ctrlLine.X2 + screen.TwipsPerPixelX, ctrlLine.Y2), COLOR_WHITE
        End If
    End If
End Sub


Public Sub Erase3DBorder(ctrlTarget As Control, nBevel As Integer, nSpace As Integer, bInset As Integer)
    Dim CTop As Integer, CLeft As Integer, CRight As Integer, CBottom As Integer
    Dim PixelX As Integer, PixelY As Integer, AddX As Integer, AddY As Integer
    Dim i As Integer

    ' Just put "No 3D" in the Tag property and your control keeps 2D
    If InStr(UCase(ctrlTarget.Tag), "NO 3D") = 0 Then
        PixelX = screen.TwipsPerPixelX
        PixelY = screen.TwipsPerPixelY
        CTop = ctrlTarget.Top - PixelY
        CLeft = ctrlTarget.Left - PixelX
        CRight = ctrlTarget.Left + ctrlTarget.Width
        CBottom = ctrlTarget.Top + ctrlTarget.Height
        If bInset Then          ' Draw border inset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                ctrlTarget.Parent.Line (CLeft - AddX, CTop - AddY)-(CRight + AddX, CTop - AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CLeft - AddX, CTop - AddY)-(CLeft - AddX, CBottom + AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CLeft - AddX, CBottom + AddY)-(CRight + AddX + PixelX, CBottom + AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CRight + AddX, CTop - AddY)-(CRight + AddX, CBottom + AddY), ctrlTarget.Parent.BackColor
            Next i
        Else                    ' Draw border outset
            For i = nSpace To (nBevel + nSpace - 1)
                AddX = i * PixelX: AddY = i * PixelY
                ctrlTarget.Parent.Line (CRight + AddX, CBottom + AddY)-(CRight + AddX, CTop - AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CRight + AddX, CBottom + AddY)-(CLeft - AddX, CBottom + AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CRight + AddX, CTop - AddY)-(CLeft - AddX - PixelX, CTop - AddY), ctrlTarget.Parent.BackColor
                ctrlTarget.Parent.Line (CLeft - AddX, CBottom + AddY)-(CLeft - AddX, CTop - AddY), ctrlTarget.Parent.BackColor
            Next i
        End If
    End If

End Sub
