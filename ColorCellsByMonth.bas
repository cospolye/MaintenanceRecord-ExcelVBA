Attribute VB_Name = "Module1"
Sub ColorCellsByMonth()
    Dim SourceSheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim CellValue As String
    Dim DateValue As Date
    Dim MonthColumn As Integer
    
    ' Set the source worksheet
    Set SourceSheet = ThisWorkbook.Sheets("ΜΑΙΝΤResults")
    
    ' Loop through the records in the source worksheet starting from row 2
    LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, "A").End(xlUp).Row
    For i = 2 To LastRow
        ' Read the date value from the "HMERAGOR" column (column D)
        DateValue = DateValueFromString(SourceSheet.Cells(i, 4).Value)
        
        ' Determine the month column based on the date
        MonthColumn = Month(DateValue) + 8 ' Offset by 8 columns (I corresponds to January)
        
        ' Check if the date is in the specified month range for the respective column
        If MonthColumn >= 9 And MonthColumn <= 20 Then ' Columns I to T
            ' Apply color based on the keyword in the "PERIGRERG" column (column E)
            CellValue = SourceSheet.Cells(i, 5).Value
            If InStr(1, CellValue, "ΣΥΝΤΗΡΗΣΗ") > 0 Then
                SourceSheet.Cells(i, MonthColumn).Interior.Color = RGB(255, 0, 0) ' Red for S??????S?
            ElseIf InStr(1, CellValue, "ΕΛΕΓΧΟΣ") > 0 Then
                SourceSheet.Cells(i, MonthColumn).Interior.Color = RGB(0, 255, 0) ' Green for ???G??S
            ElseIf InStr(1, CellValue, "ΔΙΑΚΡΙΒΩΣΗ") > 0 Then
                SourceSheet.Cells(i, MonthColumn).Interior.Color = RGB(255, 255, 0) ' Yellow for ???????OS?
            End If
        End If
    Next i
End Sub

Function DateValueFromString(DateString As String) As Date
    ' Convert a date string in the format "dd/mm/yyyy" to a Date value
    Dim DateParts() As String
    DateParts = Split(DateString, "/")
    
    If UBound(DateParts) = 2 Then
        DateValueFromString = DateSerial(CInt(DateParts(2)), CInt(DateParts(1)), CInt(DateParts(0)))
    Else
        DateValueFromString = Date
    End If
End Function

