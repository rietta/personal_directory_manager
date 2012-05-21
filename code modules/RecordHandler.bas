Attribute VB_Name = "RecordHandler"
Option Explicit
Public ItemCount As Integer
Public Sub ViewItem(itemIndex As Integer)
    ShowRecord itemIndex
End Sub
Sub ShowRecord(itemIndex As Integer)
   GetPD FreeNum, itemIndex, True
   
   If Not ViewWindowOpen Then ViewWindow.Show
   ViewWindow.Caption = Trim$(Pd.AName) & " " & Trim$(Pd.Company) & " (" & Format(CurrentIndex - 1) & " of " & Format(ItemCount) & ")"
   
   ViewWindow.lblDisplay(0) = " " & RTrim$(Pd.Company)
   ViewWindow.lblDisplay(1) = " " & RTrim$(Pd.AName)
   ViewWindow.lblDisplay(2) = " " & RTrim$(Pd.Address)
   ViewWindow.lblDisplay(3) = " " & RTrim$(Pd.City)
   ViewWindow.lblDisplay(4) = " " & RTrim$(Pd.State)
   ViewWindow.lblDisplay(5) = " " & RTrim$(Pd.Zip_Code)
   ViewWindow.lblDisplay(6) = " " & RTrim$(Pd.Home_Phone)
   ViewWindow.lblDisplay(7) = " " & RTrim$(Pd.Bus_Phone)
   ViewWindow.lblDisplay(8) = " " & RTrim$(Pd.Pager)
   ViewWindow.lblDisplay(9) = " " & RTrim$(Pd.Fax)
   ViewWindow.lblDisplay(10) = " " & RTrim$(Pd.E_Mail)
   ViewWindow.lblDisplay(11) = " " & RTrim$(Pd.WebPage)
   ViewWindow.lblDisplay(12) = " " & RTrim$(Pd.Notes)
   
    If Pd.Bookmark = False Then
        ViewWindow.cmdBookmark.Caption = "Add &Bookmark"
    Else
        ViewWindow.cmdBookmark.Caption = "Remove &Bookmark"
    End If
   
    ' Enable or Disable Prev and Next buttons
    If CurrentIndex = 2 Then
        ViewWindow.cmdPreviousItem.Enabled = False
    Else
        ViewWindow.cmdPreviousItem.Enabled = True
    End If
    If CurrentIndex = ItemCount + 1 Then
        ViewWindow.cmdNextItem.Enabled = False
    Else
        ViewWindow.cmdNextItem.Enabled = True
    End If
   
End Sub

