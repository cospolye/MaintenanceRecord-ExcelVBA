# MaintenanceRecord-ExcelVBA
Creating a Comprehensive Maintenance Records Tab from Raw Maintenance and Repair Log Files and Computing Spreadsheets
Creating a Comprehensive Maintenance Records Tab from Raw Maintenance and Repair Log Files and Computing Spreadsheets


This specific application is a VBA (Visual Basic for Applications) code for an Excel 2007 worksheet. It contains raw data covering a wide range of equipment maintenance from 2007 to 2023.


AM: The unique number of each machine (Registry Number)
SNO: Serial number of the machine
DESC: Machine model
HMERAGOR: Date of machine purchase or the date of equipment operation in the hospital
PERIGRERG: Job description
DEDATE: The date of task execution (work order date)
DVCE: Equipment type

There are three codes: DataMiningOptimized, MaintenanceCountbyYear, and ColorCellsByMonth:

DataMiningOptimized:

The VBA code performs the following tasks in Excel:

    Defines the source worksheet (SourceSheet) and the result worksheet (ResultSheet).
    Creates the result worksheet named "MAINTResults."
    Copies headers from the source worksheet (A1 to E1) to the result worksheet and adds the "DVCE" field as a header.
    Initializes the ResultRow variable to 2, which will be used for the next available row in the result worksheet.
    Iterates through the records in the source worksheet.
    Reads the value of the cell in the "PERIGRERG" column (column E) for each record.
    Checks if the text contains the keywords "ΣΥΝΤΗΡΗΣΗ" (Maintenance), "ΕΛΕΓΧΟΣ" (Inspection), or "ΔΙΑΚΡΙΒΩΣΗ" (Verification).
    If the text contains at least one of these keywords, it copies the corresponding row from the source worksheet to the result worksheet.
    Adds the value of the "DVCE" field from the source worksheet to the result worksheet for each matching record.
    After checking all records, it adjusts the sorting in the result worksheet. It sorts the results by "SNO" in ascending order, and if there are identical "SNO" values, it sorts the corresponding entries based on the "DEDATE" (date of maintenance, inspection, or verification) column in ascending order.

This is how the above code manages data and results for the desired functionality. It is optimized for performance and can quickly perform data mining on large datasets.

MaintenanceCountbyYear:

The following code is a computational code that runs in Microsoft Excel using VBA. This code takes user inputs, searches for data in a worksheet named "MAINTResults," and returns information about the number of maintenance actions for a specific machine and year. Let's describe the code in steps:

    The code begins by creating a subroutine (Sub) named "GetMaintenanceCountByYear." This subroutine executes when called by the user.

    Three variables are defined:
        "MachineSerialNumber": This variable stores the serial number of the machine entered by the user.
        "SelectedYear": This variable stores the year entered by the user.
        "MaintenanceCount": This variable is used to count the number of maintenance actions found.

    The InputBox command is used to request two inputs from the user: the "MachineSerialNumber" and the year of interest.

    Next, the code searches for data in the "MAINTResults" worksheet. It examines each row of the worksheet to find records that match both the "MachineSerialNumber" and the user-selected year.

    Each time such a record is found, the "MaintenanceCount" counter is incremented by 1.

    Finally, the results are displayed in a message box (MsgBox). If there are maintenance actions for the specified machine and year, it displays the number of maintenance actions. Otherwise, it shows a message stating that no maintenance actions were found for the specified machine and year.

ColorCellsByMonth:

    Initially, it identifies the worksheet to be processed, named "KINMHXPAR."

    It then locates the last data row in the worksheet to determine how many rows to examine.

    The code starts checking each row from the second row onwards, up to the last row.

    For each row, it reads the date from the "HMERAGOR" column (column D) and converts it into a date format.

    It then calculates the month of that date.

    Based on the calculated month, it determines which column (from I to T) corresponds to that month.

    It checks if the month falls within a specific range (e.g., if it's January).

    If the month falls within the specified range, it checks the "PERIGRERG" field (column E) to see if it contains one of the keywords "ΣΥΝΤΗΡΗΣΗ" (Maintenance), "ΕΛΕΓΧΟΣ" (Inspection), or "ΔΙΑΚΡΙΒΩΣΗ" (Verification).

    If found, it colors the corresponding cell in the respective column (e.g., if it finds "ΣΥΝΤΗΡΗΣΗ" in column E and the month is January, it will color the cell in column I).

This code adds colors to cells based on the month and the "PERIGRERG" description. This can be done for all columns from I to T in the worksheet.






