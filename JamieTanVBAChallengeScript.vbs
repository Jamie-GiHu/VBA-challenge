Sub JamieTan_VBAChallenge()

'Declare variable dimension as long rather than integer due to large dataset/high number of rows
Dim i As Long

'Declare variable dimension for ticker symbol, ticker volume, stock open price and close price
Dim ticker_symbol As Variant
Dim ticker_volume As Variant
Dim open_price As Double
Dim close_price As Double

'Declare variable dimension for each ticker symbol in the summary table
Dim Key As Integer

'Declare variable dimension as range
Dim TickerEnd As Range
Dim Rg As Range

On Error Resume Next

   
'------------------------------------------------------------------------------------'
'PART 3: LOOPING ACROSS WORKSHEET                                                    '
'------------------------------------------------------------------------------------'
    
For Each ws In Worksheets
Dim WorksheetName As String
WorksheetName = ws.Name


    '------------------------------------------------------------------'
    'PART 1: RETRIEVAL OF DATA, COLUMN CREATION & CONDITIONAL FORMATING'
    '------------------------------------------------------------------'
  
    'Give this variable a value. ticker_volume, open_price & close_price to start at 0 and Key at 2 to signify data starting at Row 2
    Key = 2
    ticker_volume = 0
    open_price = 0
    close_price = 0
    
    'Populate the column header for data output
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

    'Set the reference to our ranges being the start position and end position of our ticker
    Set TickerEnd = ws.Cells(Rows.Count, "A").End(xlUp)
    Set Rg = ws.Range("A1", TickerEnd)
     
      'Loop through all ticker symbols
      For i = 2 To UBound(Rg.Value, 1)
        
        'If Ticker symbol on next row is DIFFERENT from the current Ticker symbol, then..
        If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
            
            'Store Ticker symbol in variable
            ticker_symbol = ws.Cells(i, "A").Value
                               
                'Store closing price for the year in variable if date is on last trading day of the year for the Ticker symbol
                If ws.Cells(i, "B").Value > ws.Cells(i + 1, "B").Value Then
                
                    close_price = ws.Cells(i, "F").Value
                         
                End If
            
            'Store Ticker volume in variable cumulatively
            ticker_volume = ticker_volume + ws.Cells(i, "G").Value
            
            'Populate value stored in variable in new columns
            ws.Range("I" & Key).Value = ticker_symbol
            
            ws.Range("L" & Key).Value = ticker_volume
            
            'Calculate and populate yearly price change by Ticker in new column
            ws.Range("J" & Key).Value = (close_price - open_price)
            
                    'Conditional Formating to highlight positive change in green and negative change in red
                    If ws.Range("J" & Key).Value > 0 Then
                
                    ws.Range("J" & Key).Interior.ColorIndex = 4
                    
                    Else
                    
                    ws.Range("J" & Key).Interior.ColorIndex = 3
                    
                    End If
            
            'Calculate, populate and format Percentage Change by Ticker
            ws.Range("K" & Key).Value = Format((close_price / open_price - 1), "0.00%")
                
        'Set and reset variable before next loop
            Key = Key + 1
            
            ticker_volume = 0
            
            open_price = 0
            
            close_price = 0
        
        'If Ticker symbol on next row is THE SAME as the current Ticker symbol, then..
        Else
            
            'Store Ticker volume in variable cumulatively
            ticker_volume = ticker_volume + ws.Cells(i, "G").Value
            
            'Store opening price for the year in variable if date is on first trading day of the year for the Ticker symbol
            If ws.Cells(i, "B").Value < ws.Cells(i - 1, "B").Value Then
            
                    open_price = ws.Cells(i, "C").Value
                    
            End If
            
        End If
        
        
        
      Next i
    
    ' Autofit to display data
    ws.Columns("I:L").AutoFit
    
    
    '--------------------------------------------------------------------------'
    'PART 2: GREATEST % INCREASE, GREATEST % DECREASE AND GREATEST TOTAL VOLUME'
    '--------------------------------------------------------------------------'
    
    Dim Greatest_Increase, Greatest_Decrease As Double
    Dim Greatest_Total_Vol As Variant
    Dim x, row_GI, row_GD, row_GTV As Long
    
    'Populate the row and column headers for data output
    ws.Range("N2").Value = ("Greatest % Increase")
    ws.Range("N3").Value = ("Greatest % Decrease")
    ws.Range("N4").Value = ("Greatest Total Volume")
    ws.Range("O1:P1").Value = Array("Ticker", "Value")
    
    'Find last row in summary data from Part 1
    x = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Find and populate Ticker & Value with Greatest % Increase
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & x))
    ws.Range("P2").Value = Format(Greatest_Increase, "0.00%")
    
    row_GI = Application.WorksheetFunction.Match(Greatest_Increase, ws.Range("K2:K" & x), 0)
    ws.Range("O2").Value = ws.Cells(row_GI + 1, "I")
        
    'Find and populate Ticker & Value with Greatest % Decrease
    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & x))
    ws.Range("P3").Value = Format(Greatest_Decrease, "0.00%")
    
    row_GD = Application.WorksheetFunction.Match(Greatest_Decrease, ws.Range("K2:K" & x), 0)
    ws.Range("O3").Value = ws.Cells(row_GD + 1, "I")
    
    'Find and populate Ticker & Value with Greatest Total Volume
    Greatest_Total_Vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & x))
    ws.Range("P4").Value = Format(Greatest_Total_Vol, "General Number")
    
    row_GTV = Application.WorksheetFunction.Match(Greatest_Total_Vol, ws.Range("L2:L" & x), 0)
    ws.Range("O4").Value = ws.Cells(row_GTV + 1, "I")
    
    ' Autofit to display data
    ws.Columns("N:N").AutoFit
     
     
'Go to next worksheet
Next ws
    
    
End Sub

    
    'Populate the column header for data output
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

    'Set the reference to our ranges being the start position and end position of our ticker
    Set TickerEnd = ws.Cells(Rows.Count, "A").End(xlUp)
    Set Rg = ws.Range("A1", TickerEnd)
    Set DateEnd = ws.Cells(Rows.Count, "B").End(xlUp)
    Set DateRg = ws.Range("B2", DateEnd)
    
    'Determine first day and last day of the year
    FirstDay = Application.WorksheetFunction.Min(DateRg)
    LastDay = Application.WorksheetFunction.Max(DateRg)
     
      'Loop through all ticker symbols
      For i = 2 To UBound(Rg.Value, 1)
        
        'If Ticker symbol on next row is DIFFERENT from the current Ticker symbol, then..
        If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
            
            'Store Ticker symbol in variable
            ticker_symbol = ws.Cells(i, "A").Value
                               
                'Store closing price for the year in variable if date is on last day of the year
                If ws.Cells(i, "B").Value = LastDay Then
                
                    close_price = ws.Cells(i, "F").Value
                         
                End If
            
            'Store Ticker volume in variable cumulatively
            ticker_volume = ticker_volume + ws.Cells(i, "G").Value
            
            'Populate value stored in variable in new columns
            ws.Range("I" & Key).Value = ticker_symbol
            
            ws.Range("L" & Key).Value = ticker_volume
            
            'Calculate and populate yearly price change by Ticker in new column
            ws.Range("J" & Key).Value = (open_price - close_price)
            
                    'Conditional Formating to highlight positive change in green and negative change in red
                    If ws.Range("J" & Key).Value > 0 Then
                
                    ws.Range("J" & Key).Interior.ColorIndex = 4
                    
                    Else
                    
                    ws.Range("J" & Key).Interior.ColorIndex = 3
                    
                    End If
            
            'Calculate, populate and format Percentage Change by Ticker
            ws.Range("K" & Key).Value = Format((open_price - close_price) / open_price, "0.00%")
                
        'Set and reset variable before next loop
            Key = Key + 1
            
            ticker_volume = 0
            
            open_price = 0
            
            close_price = 0
        
        'If Ticker symbol on next row is THE SAME as the current Ticker symbol, then..
        Else
            
            'Store Ticker volume in variable cumulatively
            ticker_volume = ticker_volume + ws.Cells(i, "G").Value
            
            'Store opening price for the year in variable if date is on first day of the year
            If ws.Cells(i, "B").Value = FirstDay Then
            
                    open_price = ws.Cells(i, "C").Value
                    
            End If
            
        End If
        
        
        
      Next i
    
    ' Autofit to display data
    ws.Columns("I:L").AutoFit
    
    
    '--------------------------------------------------------------------------'
    'PART 2: GREATEST % INCREASE, GREATEST % DECREASE AND GREATEST TOTAL VOLUME'
    '--------------------------------------------------------------------------'
    
    Dim Greatest_Increase, Greatest_Decrease As Double
    Dim Greatest_Total_Vol As Variant
    Dim x, row_GI, row_GD, row_GTV As Long
    
    'Populate the row and column headers for data output
    ws.Range("N2").Value = ("Greatest % Increase")
    ws.Range("N3").Value = ("Greatest % Decrease")
    ws.Range("N4").Value = ("Greatest Total Volume")
    ws.Range("O1:P1").Value = Array("Ticker", "Value")
    
    'Find last row in summary data from Part 1
    x = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Find and populate Ticker & Value with Greatest % Increase
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & x))
    ws.Range("P2").Value = Format(Greatest_Increase, "0.00%")
    
    row_GI = Application.WorksheetFunction.Match(Greatest_Increase, ws.Range("K2:K" & x), 0)
    ws.Range("O2").Value = ws.Cells(row_GI + 1, "I")
        
    'Find and populate Ticker & Value with Greatest % Decrease
    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & x))
    ws.Range("P3").Value = Format(Greatest_Decrease, "0.00%")
    
    row_GD = Application.WorksheetFunction.Match(Greatest_Decrease, ws.Range("K2:K" & x), 0)
    ws.Range("O3").Value = ws.Cells(row_GD + 1, "I")
    
    'Find and populate Ticker & Value with Greatest Total Volume
    Greatest_Total_Vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & x))
    ws.Range("P4").Value = Format(Greatest_Total_Vol, "General Number")
    
    row_GTV = Application.WorksheetFunction.Match(Greatest_Total_Vol, ws.Range("L2:L" & x), 0)
    ws.Range("O4").Value = ws.Cells(row_GTV + 1, "I")
    
    ' Autofit to display data
    ws.Columns("N:N").AutoFit
     
     
'Go to next worksheet
Next ws
    
    
End Sub