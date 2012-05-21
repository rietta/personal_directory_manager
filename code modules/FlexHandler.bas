Attribute VB_Name = "FlexHandler"
' This module is used to power the flex data grid
' used to list names

Option Explicit

Public Sub InitializeGrid(grid As MSFlexGrid)
    grid.ColWidth(0) = 400
    grid.ColWidth(1) = (grid.Width - 400) / 2
    grid.ColWidth(2) = ((grid.Width - 400) / 2) - 400
        
    TitleHeaderRow grid
    NumberRows grid
End Sub

Public Sub NumberRows(grid As MSFlexGrid)
    Dim i As Long
    For i = 1 To grid.Rows - 1
        grid.TextMatrix(i, 0) = Format(i)   ' place string representation of number in the first column (fixed)
    Next i
End Sub

Public Sub AddRow(grid As MSFlexGrid)
    grid.Rows = grid.Rows + 1
End Sub

Public Sub LoadItemsIntoGrid(grid As MSFlexGrid)
    Dim j As Integer

    grid.Clear
    grid.Rows = 2
    
    For j = 2 To 100
        Get #FreeNum, j, Pd
        Pd.Company = Decript(Trim$(Pd.Company))
        Pd.AName = Decript(Trim$(Pd.AName))
        If RTrim$(Pd.AName) <> "" Or RTrim$(Pd.Company) <> "" Then
            If j > 2 Then AddRow grid
            grid.TextMatrix(j - 1, 0) = Format$(j - 1)
            grid.TextMatrix(j - 1, 1) = Pd.Company
            grid.TextMatrix(j - 1, 2) = Pd.AName
        End If
    Next j
    
    ItemCount = grid.Rows - 1
    
    TitleHeaderRow grid
    
    ResetPD
End Sub

Public Sub ShowSelectedRow(grid As MSFlexGrid, currentRow As Integer)
    grid.Row = currentRow
    grid.Col = 2
    grid.CellBackColor = vbHighlight
    grid.CellForeColor = vbHighlightText
    
    grid.Col = 1
    grid.CellBackColor = vbHighlight
    grid.CellForeColor = vbHighlightText
End Sub

Public Sub ShowUnselectedRow(grid As MSFlexGrid, currentRow As Integer)
    On Error Resume Next
    grid.Row = currentRow
    grid.Col = 1
    grid.CellBackColor = vbWindowBackground
    grid.CellForeColor = vbWindowText
    
    grid.Col = 2
    grid.CellBackColor = vbWindowBackground
    grid.CellForeColor = vbWindowText

End Sub

Public Sub TitleHeaderRow(grid As MSFlexGrid)
    grid.TextMatrix(0, 1) = UserField(0)    ' First field
    grid.TextMatrix(0, 2) = UserField(1)
End Sub
