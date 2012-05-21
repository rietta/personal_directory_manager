'This modual contains Proceedures to encript text strings
'The encription code was designed by Frank Rietta

'The Encript function adds the value of Cript to the ASC Value
'of the carictor and the Decript function subtracks the value.


Dim Shared Cript       'Encription Key

Function Decript (Code As String) As String
Dim i As Integer, J As Integer, Enc As String
 On Error Resume Next
 If Code = "" Then Exit Function
 'Scan String and evaluate encripted charictor
  Enc = ""
  C$ = ""
  For i = 1 To Len(Code)
     C$ = Mid$(Code, i, 1)
     Numcode = Asc(C$)
     Numcode = Numcode - Cript   'Decript Charictor with key
     Enc = Enc & Chr$(Numcode)
  Next i
  Decript = Enc
End Function

Function Encript (Code As String) As String
Dim i As Integer, J As Integer, Enc As String
 If Code = "" Then Exit Function
 'Scan String and evaluate encripted charictor
   Enc = ""
   C$ = ""
   For i = 1 To Len(Code)
      C$ = Mid$(Code, i, 1)
      Numcode = Asc(C$)
      Numcode = Numcode + Cript   'Encript Charictor with key
      Enc = Enc & Chr$(Numcode)
   Next i
   Encript = Enc
End Function

Sub SetCript (C As Integer)
  If C = 0 Then
     Cript = 10
  Else
     Cript = C
  End If
End Sub

