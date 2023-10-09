Sub SetWorkbookValues()
    'resource used in looping through worksheets'
    'https://support.microsoft.com/en-gb/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0'
    
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        'resource to solve worksheet looping issue
        'https://www.mrexcel.com/board/threads/excel-macro-looping-issue.1235764/
        sheet.Activate
    
        'Set headers and other labels
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        Call FillWorksheet
        
    Next sheet
        
    
End Sub

Sub FillWorksheet()
    'resource usedfor looping
    'https://www.excel-easy.com/vba/examples/loop-through-entire-column.html
        Dim row As Double
        Dim ticker As String
        Dim openValue As Double
        Dim closeValue As Double
        Dim volume As Double
        Dim outputRow As Double
        outputRow = 2
                        
        'Starting at 2 because header is row 1 and filled in already
        'Code for check to find last row
        'https://officetuts.net/excel/vba/count-rows-in-excel-vba/
        'I go last row + 1 because my check to create a new output row takes the current and previous row so I need the last blank row to get ZZX
        For row = 2 To Cells(Rows.Count, 1).End(xlUp).row + 1
            If Cells(row, 1).Value <> Cells(row - 1, 1).Value Then
                If Cells(row - 1, 1).Value = "<ticker>" Then
                    'For a new sheet, initialize stock Values
                    ticker = Cells(row, 1).Value
                    openValue = Cells(row, 3).Value
                    closeValue = Cells(row, 6).Value
                    volume = Cells(row, 7).Value
                Else
                    'Create a new row to output the stock values and initialize a new stock
                    Cells(outputRow, 9).Value = ticker
                    Cells(outputRow, 10).Value = closeValue - openValue
                    Call ChangeColor(outputRow)
                    Cells(outputRow, 11).Value = (closeValue - openValue) / openValue
                    'Percent Format Code
                    'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
                    Cells(outputRow, 11).NumberFormat = "0.00%"
                    Cells(outputRow, 12).Value = volume
                    outputRow = outputRow + 1
                    ticker = Cells(row, 1).Value
                    openValue = Cells(row, 3).Value
                    closeValue = Cells(row, 6).Value
                    volume = Cells(row, 7).Value
                End If
            Else
                'set the new latest close value and add to total volume
                closeValue = Cells(row, 6).Value
                volume = volume + Cells(row, 7).Value

            End If
        Next row
        
        Call OutputGreatest(outputRow)
End Sub

Sub ChangeColor(row As Double)
'color change code
'https://www.excel-easy.com/vba/examples/background-colors.html
    If Cells(row, 10) > 0 Then
        Cells(row, 10).Interior.Color = RGB(0, 255, 0)
    ElseIf Cells(row, 10) < 0 Then
        Cells(row, 10).Interior.Color = RGB(255, 0, 0)
    End If
End Sub

Sub OutputGreatest(totalRows As Double)
'Loop over output values for greatest % increase, decrease and total volume
    Dim row As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
        For row = 2 To totalRows
            If Cells(row, 11) > greatestPercentIncrease Then
                greatestPercentIncrease = Cells(row, 11)
                increaseTicker = Cells(row, 9)
            End If
            If Cells(row, 11) < greatestPercentDecrease Then
                greatestPercentDecrease = Cells(row, 11)
                decreaseTicker = Cells(row, 9)
            End If
            If Cells(row, 12) > greatestTotalVolume Then
                greatestTotalVolume = Cells(row, 12)
                volumeTicker = Cells(row, 9)
            End If
        Next row
        Cells(2, 16).Value = increaseTicker
        Cells(3, 16).Value = decreaseTicker
        Cells(4, 16).Value = volumeTicker
        Cells(2, 17).Value = greatestPercentIncrease
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).Value = greatestPercentDecrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 17).Value = greatestTotalVolume
End Sub