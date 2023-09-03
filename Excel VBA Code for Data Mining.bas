Attribute VB_Name = "Module1"
' Excel VBA Code for Data Mining - Optimized Version
' Created by Costas Polyezos'
Sub DataMiningOptimized()
    Dim SourceSheet As Worksheet
    Dim ResultSheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim CellValue As String
    Dim ResultRow As Long
    
    ' Set the source worksheet and the result worksheet
    Set SourceSheet = ThisWorkbook.Sheets("KINMHXPAR")
    Set ResultSheet = ThisWorkbook.Sheets.Add
    ResultSheet.Name = "ΜΑΙΝΤResults"
    
    ' Initialize the result row
    ResultRow = 2
    
    ' Copy headers from A1 to E1 and add "DVCE" as a header
    SourceSheet.Range("A1:E1").Copy Destination:=ResultSheet.Range("A1")
    ResultSheet.Cells(1, 7).Value = "DVCE"
    
    ' Loop through the records in the source worksheet
    LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, "A").End(xlUp).Row
    For i = 2 To LastRow
        
        ' Read the DVCE field from the source worksheet
        Dim DVCE As String
        DVCE = SourceSheet.Cells(i, 7).Value
        
        ' Read the value of the cell in the "PERIGRERG" column
        CellValue = SourceSheet.Cells(i, 5).Value
        
        ' Check if the text contains the keywords
        If InStr(1, CellValue, "ΣΥΝΤΗΡΗΣΗ") > 0 Or InStr(1, CellValue, "ΕΛΕΓΧΟΣ") > 0 Or InStr(1, CellValue, " ΔΙΑΚΡΙΒΩΣΗ") > 0 Then
            ' If any of the keywords are found, copy the row to the result worksheet
            SourceSheet.Cells(i, 1).Resize(1, 5).Copy Destination:=ResultSheet.Cells(ResultRow, 1)
            ResultSheet.Cells(ResultRow, 7).Value = DVCE ' Add the DVCE field
            ResultRow = ResultRow + 1
        End If
    Next i
    
    ' Adjust sorting in the result worksheet
    ResultSheet.Sort.SortFields.Clear
    ResultSheet.Sort.SortFields.Add Key:=ResultSheet.Range("B2:B" & ResultRow - 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ResultSheet.Sort.SortFields.Add Key:=ResultSheet.Range("E2:E" & ResultRow - 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ResultSheet.Sort
        .SetRange ResultSheet.Range("A1:G" & ResultRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub



