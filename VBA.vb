Sub StockAnalysis()

Dim ws As Worksheet

Dim Stock_Volume As Long
Dim Ticker_code As String
Dim Year As String
Dim NumRows As Long
Dim TickerRow As Long
Dim TickerStart As Double
Dim TickerEnd As Double
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

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
  
    NumRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_code = ""
    TickerRow = 1
    
    'Looking through rows
    For i = 2 To NumRows
        Stock_Volume = ws.Cells(i, 7)
    
        'Looking for new ticker code
        If Ticker_code <> ws.Cells(i, 1) Then
            TickerStart = ws.Cells(i, 3) 'Captures Ticker open amount
            TickerRow = TickerRow + 1
            
            ws.Cells(TickerRow, 9) = ws.Cells(i, 1) 'Writes the Ticker code
            ws.Cells(TickerRow, 12) = Stock_Volume 'Writes the Stock Volume
        
        'Looking for existing ticker code
        Else
            TickerEnd = ws.Cells(i, 6) 'Captures Ticker close amount
            
            ws.Cells(TickerRow, 10) = TickerEnd - TickerStart 'Calculates Yearly Change
            
            'Conditional formatting to turn positives into Green and negatives into Red
            If ws.Cells(TickerRow, 10) < 0 Then
                ws.Cells(TickerRow, 10).Interior.Color = vbRed
            ElseIf ws.Cells(TickerRow, 10) > 0 Then
                ws.Cells(TickerRow, 10).Interior.Color = vbGreen
            End If
            
            'To fix debug error
            If TickerStart = 0 Then
                ws.Cells(TickerRow, 11) = 0
            Else
                ws.Cells(TickerRow, 11) = ws.Cells(TickerRow, 10) / TickerStart 'Calculates Percentage Change
            End If
            
            ws.Cells(TickerRow, 12) = ws.Cells(TickerRow, 12) + Stock_Volume 'Adds up the Total Stock Volume
            ws.Cells(TickerRow, 11).NumberFormat = "0.00%" 'Converts to percent format
            
        End If
        
        Ticker_code = ws.Cells(i, 1)
        
    Next i

'Range in which to find the smallest and largest percent change value
lastRow = ws.Range("K1").End(xlDown).Row
strData = "K2:K" & lastRow & ""
strData2 = "L2:L" & lastRow & ""

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