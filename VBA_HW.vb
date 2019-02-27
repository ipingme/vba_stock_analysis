Sub Stock_analysis()

    Dim stock_name As String
    Dim stock_total As Double
    Dim LastRow As Long
    Dim stock_open As Double
    Dim stock_close As Double
    Dim yearly_change As Double
    Dim summary_table_row As Integer
    Dim strData As String
    Dim strData2 As String
    Dim rng As Range
    Dim rng2 As Range
    Dim vValue As Variant
    Dim wValue As Variant
    Dim mValue As Variant
    Dim rngCol As Range
    Dim lngRow As Long
    Dim maxRow As Long
    Dim maxRow2 As Long
    Dim maxrngAdd As Range
    Dim maxrngAdd2 As Range
    Dim rngAdd As Range
    
    For Each ws In Worksheets
        'ws.Activate

        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'To get last row
        
        stock_total = 0
        summary_table_row = 2
        check_open = ""
        check_close = ""
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stock_name = ws.Cells(i, 1).Value
                If check_close = "" Then
                    stock_close = ws.Range("F" & i)
                    check_close = "y" 'This is to make sure stock close does not get overridden
                    If stock_open <> 0 Then 'This is to account for any errors with dividing by 0
                        ws.Range("J" & summary_table_row).Value = stock_close - stock_open
                        ws.Range("K" & summary_table_row).Value = (stock_close - stock_open) / stock_open
                    Else
                        ws.Range("J" & summary_table_row).Value = stock_close - stock_open
                        ws.Range("K" & summary_table_row).Value = 0
                    End If
                    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"  'Percent formatting on Percent change
                    
                    'This is to reset the ticker open and close values
                    check_close = ""
                    check_open = ""
                End If
                
                stock_total = stock_total + ws.Cells(i, 7).Value
                ws.Range("I" & summary_table_row).Value = stock_name 'Write stock name
                ws.Range("L" & summary_table_row).Value = stock_total 'Write stock total volume
                
                'Conditional formatting for Green and Red colors
                If ws.Range("J" & summary_table_row) >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4 'Green
                ElseIf ws.Range("J" & summary_table_row) < 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3 'Red
                End If
                
                'This is to move to the next row in the summary table
                summary_table_row = summary_table_row + 1
                stock_total = 0
            Else
                If check_open = "" Then
                    stock_open = ws.Range("C" & i)
                    check_open = "y" 'This is to make sure stock_open does not get overridden
                End If
                stock_total = stock_total + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Range in which to find the smallest and largest percent change value
        LastRow = ws.Range("K1").End(xlDown).Row
        strData = "K2:K" & LastRow & ""
        strData2 = "L2:L" & LastRow & ""
        
        Set rng = ws.Range(strData)
        Set rng2 = ws.Range(strData2)
        
        'Determines smallest value in range
        vValue = Application.WorksheetFunction.Min(rng)
        wValue = Application.WorksheetFunction.Max(rng)
        mValue = Application.WorksheetFunction.Max(rng2)
        
            For Each rngCol In rng.Columns
            
                'Determines in case the smallest value exists in a particular column
                If Application.WorksheetFunction.CountIf(rngCol, vValue) > 0 Then
                
                    'Returns row number of the smallest and largest value, in the column which has the same
                    lngRow = Application.WorksheetFunction.Match(vValue, rngCol, 0)
                    maxRow = Application.WorksheetFunction.Match(wValue, rngCol, 0)
        
                    'Returns cell address of the smallest value
                    Set rngAdd = rngCol.Cells(lngRow, 1)
                    Set maxrngAdd = rngCol.Cells(maxRow, 1)
                               
                    'Message displays the searched range, smallest value, and its address
                    
                    'MsgBox "Smallest Value in Range(""" & strData & """) is " & vValue & ", in Cell " & rngAdd.Address & "."
                    'MsgBox "Largest Value in Range(""" & strData & """) is " & wValue & ", in Cell " & maxrngAdd.Address & "."
                    
                    ws.Range("P2") = ws.Cells(maxRow + 1, 9)
                    ws.Range("Q2") = ws.Cells(maxRow + 1, 11)
                    ws.Range("P3") = ws.Cells(lngRow + 1, 9)
                    ws.Range("Q3") = ws.Cells(lngRow + 1, 11)
                    
                    ws.Cells(2, 17).NumberFormat = "0.00%" 'Converts to percent format
                    ws.Cells(3, 17).NumberFormat = "0.00%" 'Converts to percent format
        
                    Exit For
        
                End If
            
            Next
            
            For Each rngCol In rng2.Columns
            
                'Determines in case the smallest value exists in a particular column
                If Application.WorksheetFunction.CountIf(rngCol, mValue) > 0 Then
                
                    'Returns row number of the smallest and largest value, in the column which has the same
                    maxRow2 = Application.WorksheetFunction.Match(mValue, rngCol, 0)
        
                    'Returns cell address of the smallest value
                    Set maxrngAdd2 = rngCol.Cells(maxRow2, 1)
                    
                    'MsgBox "Largest Value in Range(""" & strData2 & """) is " & mValue & ", in Cell " & maxrngAdd2.Address & "."
                    
                    ws.Range("P4") = ws.Cells(maxRow2 + 1, 9)
                    ws.Range("Q4") = ws.Cells(maxRow2 + 1, 12)
                    ws.Columns("O").AutoFit
                    ws.Columns("J:L").AutoFit
                    
                    Exit For
                
                End If
            
            Next
        
    Next ws

End Sub

