Sub StockVolume()
'This code is a script that will loop through one year of stock data for each run and 
'return the total volume each stock had over that year. I could only get this code to
'run on the test excel file. Even then I couldn't get it to account of the year. Reoccurring
'issue with getting Overflow error and Overflow error 6.

'This section is for my variables
    Dim Ticker_Current As String
    Dim Count_Ticker_Type As Integer
    Dim Summary_Ticker_Total As Integer
    Dim Summary_Ticker_Row As Integer
 '   Dim Last_Row As Integer
 '   Dim New_Row As Integer
 '   Dim Current_Row As Integer

    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    Count_Ticker_Type = 0
    Summary_Ticker_Row = 2
    Summary_Ticker_Total = 0
'    newRow = 2
'    currentRow = 2
'    Cells(newRow, 10) = tickerCurrent

    For i = 2 To Last_Row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Make current stock ticker value the equal to currect 1st column cell
            Ticker_Current = Cells(i, 1).Value
            
            'Set current count of stock ticker value to total count
            Summary_Ticker_Total = Count_Ticker_Type + 1
            
            'Set up summary table
            Range("J" & Summary_Ticker_Row).Value = Ticker_Current
            Range("K" & Summary_Ticker_Row).Value = Summary_Ticker_Total
            
            'Set new summary row for next stock
            Summary_Ticker_Row = Summary_Ticker_Row + 1
            Count_Ticker_Type = 0
        Else
            Count_Ticker_Type = Count_Ticker_Type + 1
        End If
    
    Next i

End Sub
