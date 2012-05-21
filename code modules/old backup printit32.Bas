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
    Printer.ScaleWidth = Printer.ScaleWidth - 1.5
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
    Dim i As Integer
    
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
    
    If Printer.TextWidth(StringToPrint) > Printer.ScaleWidth - nCenterLine Then
        PrintedLine$ = ""
        NextWord$ = GetNextWord(StringToPrint)
        
        Do While NextWord$ <> ""
            temp$ = PrintedLine$ + NextWord$
            AddMargin nPrintMargin
            'Check the width of the temp string
            If Printer.TextWidth(temp$) > Printer.ScaleWidth - nCenterLine Then
                'Make sure the line will fit vertically on the page
                If Printer.CurrentY > Printer.ScaleHeight - Printer.TextHeight(PrintedLine$) Then
                    Printer.NewPage
                End If
                PrintNextLine PrintedLine$    'Send the line to the printer
                PrintedLine$ = NextWord$      'Start a new line to print
            Else
                PrintedLine$ = temp$
            End If
            NextWord$ = GetNextWord(StringToPrint)
        Loop
        PrintNextLine PrintedLine$ + vbCrLf   'Print remainder
    Else
        'The string is fine as it is so just print it.
        'Make sure the line will fit on the page
        AddMargin nPrintMargin
        If Printer.CurrentY > Printer.ScaleHeight - Printer.TextHeight(StringToPrint) Then
            Printer.NewPage
            For i% = 1 To 4
               Printer.Print
            Next i%
        
        End If
        PrintNextLine StringToPrint + vbCrLf
        
    End If
        

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
