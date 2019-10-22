Sub DeleteRows()
Dim fullcounter As Long
fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
For i = 2 To fullcounter
If Not Worksheets("Tier 2").Cells(i, 3).Value = "Tier 2" Then
Worksheets("Tier 2").Cells(i, 3).EntireRow.Interior.ColorIndex = 6
End If
Next i

End Sub
