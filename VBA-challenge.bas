Attribute VB_Name = "Module1"
Sub stockdata()
    Dim i As LongLong
    Dim Current_Cell As String
    Dim Future_Cell As String
    Dim symbol As LongLong
    Dim Change As Double
    Dim row_count As LongLong
    Dim Tik_open As Double
    Dim Tik_close As Double
    Dim Previous_Cell As String
    Dim j As LongLong
    Dim yearlychange As Double
    Dim Percentchange As Double
    Dim TSV As LongLong
    Dim Lastrow As LongLong
    Dim ws_name As String

    
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
For Each ws In Worksheets
    ws_name = ws.Name
    ws.Cells(1, 10).Value = "ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    TSV = 0
    symbol = 2
'grab first open ticker price before we start looping
    Tik_open = ws.Cells(2, 3).Value
    For i = 2 To Lastrow
        'setting the values to determine the change in ticker name
        Current_Cell = ws.Cells(i, 1).Value
        Future_Cell = ws.Cells(i + 1, 1).Value
    
        'Compiling total Stock Volume
            TSV = TSV + ws.Cells(i, 7).Value
        
        'saying if there's a difference to print the current ticker name in the J column
        If Current_Cell <> Future_Cell Then
            ws.Cells(symbol, 10).Value = Current_Cell
            
        'Taking the values from the closing column of the last Ticker symbol and storing it as Tik_close
           Tik_close = ws.Cells(i, 6).Value
           Change = Tik_close - Tik_open
           ws.Cells(symbol, 11).Value = Change
            
            'determine the yearly percent change per ticker
                If Tik_open = 0 Then
                    ws.Cells(symbol, 12).Value = " New Stock"
                    
                Else
                    Percentchange = Change / Tik_open
                    ws.Cells(symbol, 12).Value = Percentchange
                    ws.Cells(symbol, 12) = FormatPercent(ws.Cells(symbol, 12), 2)
                End If
        
        'taking values from first opening of a ticker and storing it as a variable Tik_open
            Tik_open = ws.Cells(i + 1, 3).Value
            
            
            
           'Formating cells to be red if negative change and green if positive change
                If ws.Cells(symbol, 11).Value > 0 Then
                    ws.Cells(symbol, 11).Interior.ColorIndex = 4
        
                Else
                    ws.Cells(symbol, 11).Interior.ColorIndex = 3
                
                End If
            

            'Assigning Total stock volume to cell for display
            ws.Cells(symbol, 13).Value = TSV
            
            'Going to the next symbol
            symbol = symbol + 1
            'Restarting the Total stock volume to 0 for the next ticker
            TSV = 0
          End If
       Next i
    Next ws
End Sub

