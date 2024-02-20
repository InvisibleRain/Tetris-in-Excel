'This module contains LibreOffice Basic subroutine to replace broken VBA functions
Option VBASupport 0

Sub TestAdd()
	Call Basic.AddBordersToCell(4,5)
End Sub


Sub TestRemove()
	Call Basic.RemoveBordersFromCell(4,5)
End Sub

Public Sub AddBordersToCell(row As Integer, col As Integer)
	Dim Doc As Object
	Dim Sheet As Object
    Dim Cell As Object
    Dim oLine As New com.sun.star.table.BorderLine2

    With oLine
        .Color = 0
        .LineStyle = com.sun.star.table.BorderLineStyle.SOLID
        .LineWidth = 23
    End With
	
	Doc = ThisComponent
    Sheet = Doc.Sheets.getByName("Sheet1")
    'Unlike in VBA, the first column has the index 0 and not the index 1.
    Cell = Sheet.getCellByPosition(col - 1, row - 1)

    With Cell
        .LeftBorder = oLine
        .RightBorder = oLine
        .TopBorder = oLine
        .BottomBorder = oLine
    End With
End Sub

Public Sub RemoveBordersFromCell(row As Integer, col As Integer)
	Dim Doc As Object
	Dim Sheet As Object
    Dim Cell As Object
    Dim oLine As New com.sun.star.table.BorderLine2

    With oLine
        .LineStyle = com.sun.star.table.BorderLineStyle.NONE
    End With
	
	Doc = ThisComponent
    Sheet = Doc.Sheets.getByName("Sheet1")
    'Unlike in VBA, the first column has the index 0 and not the index 1.
    Cell = Sheet.getCellByPosition(col - 1, row - 1)

    With Cell
        .LeftBorder = oLine
        .RightBorder = oLine
        .TopBorder = oLine
        .BottomBorder = oLine
    End With
End Sub
