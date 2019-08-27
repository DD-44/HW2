Sub Year_Analysis()

'Variables for the EASY & MODERATE part
Dim WS As Worksheet
Dim LastRow As Long
Dim FirstTickerRow As Long
Dim VolumeSum As Double
Dim YearlyChange As Double
Dim PercentChange As Variant
Dim StoreRowIndex As Long
Dim i As Long
Dim j As Long

'loop to go through all worksheets
For Each WS In ThisWorkbook.Worksheets

'Make the worksheet active (in case many sheets are selected before running the macro)
WS.Activate

'Get the LastRow of the active worksheet
LastRow = ActiveSheet.UsedRange.Rows.Count

'Sort column "A", and using column "B" as a secondary sort reference
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range(Cells(1, 1), Cells(LastRow, 7))
     .Header = xlYes
     .Apply
End With

'Define the row where the results will be stored
StoreRowIndex = 1

'Define the new headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yeary Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Loop through each row
For i = 2 To LastRow
    
    'Check the previous ticker and current ticker
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    StoreRowIndex = StoreRowIndex + 1
    Cells(StoreRowIndex, 9) = Cells(i, 1).Value
    VolumeSum = Cells(i, 7).Value
    Cells(StoreRowIndex, 12) = VolumeSum 'store the new volume sum in the correct cell

    ElseIf Cells(i - 1, 1).Value = Cells(i, 1).Value Then
    VolumeSum = VolumeSum + Cells(i, 7).Value
    Cells(StoreRowIndex, 12) = VolumeSum 'store the new volume sum in the correct cell
    
    End If
Next i


'-------------------------------------------
'  MODERATE
'-------------------------------------------

'Row index for the grouped ticker values (columns I to L)
StoreRowIndex = 1

For j = 2 To LastRow

    'if to go though the ticker names. If the previous ticker is different than the current one...
    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
    StoreRowIndex = StoreRowIndex + 1
    FirstTickerRow = j 'store the value of the First row of the group of tickers
    YearlyChange = Cells(j, 6).Value - Cells(j, 3).Value '
    'IF to bypass zero value in the denominator
    If Cells(j, 3).Value = 0 Then
    PercentChange = 0
    Else
    PercentChange = Cells(j, 6).Value / Cells(j, 3).Value
    End If
    
    Cells(StoreRowIndex, 10) = YearlyChange
        'If to change the color based on the value
        If Cells(StoreRowIndex, 10) > 0 Then
        Cells(StoreRowIndex, 10).Interior.ColorIndex = 4
        ElseIf Cells(StoreRowIndex, 10) < 0 Then
        Cells(StoreRowIndex, 10).Interior.ColorIndex = 3
        End If
        
    Cells(StoreRowIndex, 11) = PercentChange
    
    'if the previous ticker is equal the current one...
    ElseIf Cells(j - 1, 1).Value = Cells(j, 1).Value Then
    YearlyChange = Cells(j, 5).Value - Cells(FirstTickerRow, 3).Value
        'IF to bypass zero value in the denominator
        If Cells(FirstTickerRow, 3).Value = 0 Then
        PercentChange = 0
        Else
        PercentChange = Cells(j, 5).Value / Cells(FirstTickerRow, 3).Value
        End If
        
    Cells(StoreRowIndex, 10) = YearlyChange
        If Cells(StoreRowIndex, 10) > 0 Then
        Cells(StoreRowIndex, 10).Interior.ColorIndex = 4
        ElseIf Cells(StoreRowIndex, 10) < 0 Then
        Cells(StoreRowIndex, 10).Interior.ColorIndex = 3
        End If
    Cells(StoreRowIndex, 11) = PercentChange
    
    End If
Next j

'-------------------------------------------
'  HARD
'-------------------------------------------

'variable definition
Dim k As Long
Dim LastRow2 As Long
Dim GreatestValue As Double
Dim GVName As String
Dim LowestValues As Double
Dim LVName As String
Dim GreatestVolume As Double
Dim GVolName As String

'Formatting the additional fiels needed
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Checking the last row based on the data added in the MODERATE part
LastRow2 = WS.Cells(WS.Rows.Count, "K").End(xlUp).Row

'Loop and IFs to checking for the values requested in the HARD part
For k = 2 To LastRow2

    If k = 2 Then
    GreatestValue = Cells(k, 11).Value
    LowestValue = Cells(k, 11).Value
    GreatestVolume = Cells(k, 12).Value
    GVName = Cells(k, 9).Value 'Ticker Name for the Greatest Value
    LVName = Cells(k, 9).Value 'Ticker Name for the Lowest Value
    GVolName = Cells(k, 9).Value 'Ticker Name for the Greatest Volume
    
    Else
    
        If GreatestValue < Cells(k, 11).Value Then
        GreatestValue = Cells(k, 11).Value
        GVName = Cells(k, 9).Value
        End If
        
        If LowestValue > Cells(k, 11).Value Then
        LVName = Cells(k, 9).Value
        End If
        
        If GreatestVolume < Cells(k, 12).Value Then
        GVolName = Cells(k, 9).Value
        End If
        
    End If

Next k

'Store the values and names in the cells
    Cells(2, 17).Value = GreatestValue
    Cells(2, 16).Value = GVName
    Cells(3, 17).Value = LowestValue
    Cells(3, 16).Value = LVName
    Cells(4, 17).Value = GreatestVolume
    Cells(4, 16).Value = GVolName


'AutoFit columns
    ActiveSheet.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
'Freeze panes to always show the header
    ActiveWindow.FreezePanes = False
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True

'Set the header font to Bold
    LastColumn = ActiveSheet.UsedRange.Columns.Count
    Range("A1", Cells(1, LastColumn)).Font.Bold = True
    Range("A1").Select

Next

End Sub