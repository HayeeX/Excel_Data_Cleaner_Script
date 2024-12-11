Sub DataCleaningUtility()

    ' Declare variables for the worksheet, data range, and other components
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim cell As Range
    Dim avg As Double, stdDev As Double
    
    ' Set the active worksheet (you can modify this to target specific sheets)
    Set ws = ActiveSheet

    ' Identify the last row and column of data in the active worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Define the data range to clean
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' --- Step 1: Remove Duplicates ---
    ' Remove duplicate rows based on the first two columns (adjust as needed)
    On Error Resume Next
    rng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    On Error GoTo 0

    ' --- Step 2: Handle Missing Values ---
    ' Replace blank cells with "N/A" to ensure no empty cells remain
    On Error Resume Next
    rng.SpecialCells(xlCellTypeBlanks).Value = "N/A"
    On Error GoTo 0

    ' --- Step 3: Standardize Text in Column A ---
    ' Trim whitespace and convert text in column A to uppercase
    For Each cell In rng.Columns(1).Cells
        If Not IsEmpty(cell.Value) Then
            cell.Value = UCase(Trim(cell.Value)) ' Convert to uppercase and remove extra spaces
        End If
    Next cell

    ' --- Step 4: Remove Outliers in Column B ---
    ' Calculate the average and standard deviation for numerical data in column B
    On Error Resume Next
    avg = WorksheetFunction.Average(rng.Columns(2))
    stdDev = WorksheetFunction.StDev(rng.Columns(2))
    On Error GoTo 0

    ' Loop through column B and clear contents of outliers
    For Each cell In rng.Columns(2).Cells
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            If Abs(cell.Value - avg) > 2 * stdDev Then
                cell.ClearContents ' Remove outlier value
            End If
        End If
    Next cell

    ' --- Step 5: Format Dates in Column C ---
    ' Apply consistent date formatting to column C
    On Error Resume Next
    rng.Columns(3).NumberFormat = "mm/dd/yyyy"
    On Error GoTo 0

    ' --- Step 6: Provide Feedback to the User ---
    ' Inform the user that the cleaning process is complete
    MsgBox "Data Cleaning Complete! Your dataset has been cleaned and organized.", vbInformation, "Data Cleaning Utility"

End Sub
