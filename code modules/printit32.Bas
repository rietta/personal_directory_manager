Attribute VB_Name = "PrintIt_32bit"
Option Explicit
Public linesPrinted As Integer
Public nPrintMargin As Integer
Public printHexColor As Long
Public nPrintColor As Integer

Dim nColOffset As Integer, nCenterLine As Integer
Dim nMarginOffset As Integer, TopY As Integer



'  Sets the printer to print a margin specified
'  the printM parameter.
'  5 = One Inch.
'  6 = One Millimeter (default margin)
'  7 = One Centimeter
Sub AddMargin(printM As Integer)
 If printM <> 0 Then
    Printer.ScaleMode = printM
    'Printer.ScaleWidth = Printer.ScaleWidth - 1.5
    Printer.CurrentX = nColOffset  ' 1
 End If
End Sub

Function GetNextWord(StringToPrint) As String
    Dim SpacePos
    'This function finds and returns the next word in a string
    
    'Find the position of the first space character
    SpacePos = InStr(StringToPrint, " ")

    If SpacePos = 0 Then
        'If there are no spaces, return the entire string
        GetNextWord = StringToPrint
        StringToPrint = ""
    Else
        'Return the word
        GetNextWord = Left$(StringToPrint, SpacePos)
        'Eliminate the returned word from the string
        StringToPrint = right$(StringToPrint, Len(StringToPrint) - SpacePos)
    End If

End Function

Sub PrintLine(StringToPrint As String, Optional nLeftX, Optional nCntrCol)
    
    Dim PrintedLine As String, NextWord As String, temp As String
    Dim i As Integer, lne As Integer, nLineOver As Integer, offset As Integer
    
    If nCntrCol > 0 Then
        nCenterLine = nCntrCol
    Else
        nCenterLine = 0
    End If
    
    If nLeftX > 0 Then
        nColOffset = nLeftX
    Else
        nColOffset = 1
    End If
    
    'Check if the line is to long since maybe it doesn't even need word-wrapped
    ' The nCenterLine varible defines the offset of a column
    StringToPrint = Trim$(StringToPrint)
    
    ' Create an Array of Lines That We Know,
    ' meaning that a CR or LR are specified.
    ReDim LineBuff(0) As String
    i = 0
    Do
        offset = InStr(1, StringToPrint, vbCrLf)
    
        If offset > 0 Then
            ReDim Preserve LineBuff(i) As String
            LineBuff(i) = Left(StringToPrint, offset - 1)
            StringToPrint = Mid(StringToPrint, offset + 2)
            i = i + 1
        Else
            If i = 0 Then
                LineBuff(i) = StringToPrint
            Else
                ReDim Preserve LineBuff(i) As String
                LineBuff(i) = StringToPrint
            End If
            Exit Do
        End If
    Loop
    
        
    ' Loop Through the Array and Print Each Line
    
    For lne = 0 To UBound(LineBuff)
        
        StringToPrint = LineBuff(i)
    
        If Printer.TextWidth(StringToPrint) > (Printer.ScaleWidth - nCenterLine) Then
        
            ' Print to the end of the line and wrap
            ' using a dash
            For i = Len(StringToPrint) To 1 Step -1
                PrintedLine = Left$(StringToPrint, i)
                If (Printer.ScaleWidth - nCenterLine) > Printer.TextWidth(PrintedLine) Then
                    PrintedLine = Left$(StringToPrint, i - 1)
                    AddMargin nPrintMargin
                    Printer.Print PrintedLine; " - "
                    StringToPrint = Mid$(StringToPrint, i)
                    Exit For
                End If
            Next i
            
            AddMargin nPrintMargin
            For i = 1 To Len(StringToPrint)
                Printer.Print Mid$(StringToPrint, i, 1);
            Next i
            
            Printer.Print
        Else
            AddMargin nPrintMargin
            Printer.Print StringToPrint
        End If
    
    
    Next lne
    ' Reset Values to Full Page
    nCenterLine = 0
    nColOffset = 1
  
End Sub

Sub PrintNextLine(StringToPrint As String)
Dim text As String, Location As Integer, temp As String
Dim i As Integer
For i = 1 To Len(StringToPrint)
  If Mid$(StringToPrint, i, 1) = Chr$(13) Then
     AddMargin nPrintMargin
     If Left$(text, 1) = Chr$(10) Then
         Printer.Print right$(text, Len(text) - 1)  'Don't print linefeed
         linesPrinted = linesPrinted + 1
     Else
         Printer.Print text
         linesPrinted = linesPrinted + 1
     End If
     text = ""
  Else
     text = text & Mid$(StringToPrint, i, 1)
  End If
Next i

End Sub

Sub PrintText(StringToPrint As String)
Dim text As String, i As Integer

If InStr(StringToPrint, Chr(13)) = 0 Then StringToPrint = StringToPrint + vbCrLf

For i = 1 To Len(StringToPrint)
  If Mid$(StringToPrint, i, 1) = Chr$(13) Then
     If Left$(text, 1) = Chr$(10) Then
         PrintLine right$(text, Len(text) - 1)  'Don't print linefeed
     Else
         PrintLine text
     End If
     text = ""
  Else
     text = text & Mid$(StringToPrint, i, 1)
  End If
Next i

End Sub


Public Sub PrintCenterText(text As String)
    Dim textW As Integer
    textW = Printer.TextWidth(text)
    
    Printer.CurrentX = ((Printer.ScaleWidth - nCenterLine) / 2) - (textW / 2)
    Printer.Print text
End Sub

Public Sub AdjustPrintItHandlers(nCol, nCL, nMO, nTY)
    nColOffset = nCol
    nCenterLine = nCL
    nMarginOffset = nMO
    TopY = nTY
End Sub
