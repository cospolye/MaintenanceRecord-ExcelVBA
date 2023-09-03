Attribute VB_Name = "Module1"
' Excel VBA Code to Retrieve Maintenance Count by Year
' Created by Costas Polyezos

Sub GetMaintenanceCountByYear()
    Dim MachineSerialNumber As String
    Dim SelectedYear As Integer
    Dim MaintenanceCount As Integer
    
    ' Prompt the user for the MACHINE SERIAL NUMBER
    MachineSerialNumber = InputBox("Enter the MACHINE SERIAL NUMBER:")
    
    ' Prompt the user for the year
    SelectedYear = InputBox("Enter the year for which you want to find the number of maintenances:")
    
    ' Search for data in the worksheet
    Dim LastRow As Long
    LastRow = Sheets("ÌÁÉÍÔResults").Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To LastRow
        If Sheets("ÌÁÉÍÔResults").Cells(i, 2).Value = MachineSerialNumber And Year(Sheets("ÌÁÉÍÔResults").Cells(i, 4).Value) = SelectedYear Then
            MaintenanceCount = MaintenanceCount + 1
        End If
    Next i
    
    ' Display the results
    If MaintenanceCount > 0 Then
        MsgBox "For MACHINE SERIAL NUMBER: " & MachineSerialNumber & " and the year " & SelectedYear & vbCrLf & _
               "Number of Maintenances for the specified year: " & MaintenanceCount, vbInformation, "Maintenance Information"
    Else
        MsgBox "No maintenances were found for MACHINE SERIAL NUMBER: " & MachineSerialNumber & " and the year " & SelectedYear, vbExclamation, "Error"
    End If
End Sub


