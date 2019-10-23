Sub CopyTier2Data()
Worksheets("Raw").Copy after:=Worksheets(Sheets.Count)
ActiveSheet.Name = "Tier 2"
End Sub


Sub DeleteRows()

Application.ScreenUpdating = False

Dim fullcounter As Integer
    
    fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
    For i = fullcounter To 2 Step -1
    If Not Sheets("Tier 2").Cells(i, 3).Value = "Tier 2" Then
    Worksheets("Tier 2").Cells(i, 3).EntireRow.Delete
    End If
    Next i

    fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
    For i = fullcounter To 2 Step -1
    If Sheets("Tier 2").Cells(i, 3).Value = "Tier 2" And _
    Sheets("Tier 2").Cells(i, 6).Value = "Local Newspaper" Then
    Worksheets("Tier 2").Cells(i, 3).EntireRow.Delete
    End If
    Next i

    fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
    For i = fullcounter To 2 Step -1
    If Sheets("Tier 2").Cells(i, 3).Value = "Tier 2" And _
    Sheets("Tier 2").Cells(i, 6).Value = "Magazines" Then
    Worksheets("Tier 2").Cells(i, 3).EntireRow.Delete
    End If
    Next i

    fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
    For i = fullcounter To 2 Step -1
    If Sheets("Tier 2").Cells(i, 3).Value = "Tier 2" And _
    Sheets("Tier 2").Cells(i, 6).Value = "Magazines" Then
    Worksheets("Tier 2").Cells(i, 3).EntireRow.Delete
    End If
    Next i

    fullcounter = Worksheets("Tier 2").Cells(Rows.Count, 3).End(xlUp).Row
    For i = fullcounter To 2 Step -1
    If Sheets("Tier 2").Cells(i, 3).Value = "Tier 2" And _
    Sheets("Tier 2").Cells(i, 6).Value = "OOH" Then
    Worksheets("Tier 2").Cells(i, 3).EntireRow.Delete
    End If
    Next i


On Error Resume Next
Application.ScreenUpdating = True

MsgBox "Succeeded."

End Sub
