Private Sub Workbook_Open()
    Call SetScrollArea
End Sub

Sub SetScrollArea()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.ScrollArea = "A1"
    Next ws
End Sub
