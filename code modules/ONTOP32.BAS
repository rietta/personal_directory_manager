Attribute VB_Name = "ONTOP"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
'Declare Function CreateBubble% Lib "bubble.dll" (ByVal X%, ByVal Y%, ByVal xs%, ByVal ys%, ByVal title$, ByVal txt$)
'Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)

Sub NoStayOnTop(frm As Form)
Dim success As Long
success = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub StayOnTop(frm As Form)
Dim success As Long
success = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

