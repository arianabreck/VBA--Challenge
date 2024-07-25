# VBA--Challenge
Sub ticker()


        
       'Set a variable for holding the ticker name
            Dim ticker_name As String
        'Set a varable for keeping count on total volue
        Dim ticker_volume As Double
        ticker_volume = 0
        'Keep track of the location for tickername
        Dim Ticker_row As Integer
        Ticker_row = 2
        
         'Set initial open_price
        Dim Open_price As Double
      
        Open_price = Cells(2, 3).Value
        Dim Close_price As Double
        Dim Year_change As Double
        Dim Percent_hange As Double
        'Label Table headers
                         Cells(1, 9).Value = "Ticker"
                        Cells(1, 10).Value = "Quarterly Change"
                        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'Count rows in first column
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        
          'Loop through each ticker name
                 For i = 2 To Last_Row
            'Look to see if current cell is like the next one
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              
              
       'Set the ticker name
              ticker_name = Cells(i, 1).Value
              
            
       'Add volume of trade
              ticker_volume = ticker_volume + Cells(i, 7).Value
              
           'Output tickername in table
              Range("I" & Ticker_row).Value = ticker_name
              
         'Output trade volume for tickers
              Range("L" & Ticker_row).Value = ticker_volume
              
        'Calculate change in price
                    Close_price = Cells(i, 6).Value
              
              'Calculate quarterly change
              Quarterly_change = (Close_price - Open_price)
              
       'Output Quarterly change in table
              Range("J" & Ticker_row).Value = Quarterly_change
             
                If (Open_price = 0) Then
                    percent_change = 0
                Else
                    percent_change = Quarterly_change / Open_price
                End If
         'Output the Percentage change for each ticker
              Range("K" & Ticker_row).Value = percent_change
              Range("K" & Ticker_row).NumberFormat = "0.00%"
              
              
              'Restart tand add to row counter
             Ticker_row = Ticker_row + 1
           'Restart volume
              ticker_volume = 0
              
              
              'Restart open price
              Open_price = Cells(i + 1, 3)
            Else
               'Sum the volume of trade
              ticker_volume = ticker_volume + Cells(i, 7).Value
           End If
        Next i
        
        
    'Color positive change in green and negative change in Red
                     last_row_Ticker_name = Cells(Rows.Count, 9).End(xlUp).Row
    'Color code quarterly change
    For i = 2 To last_row_Ticker_name
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
        End If
     Next i
     
     End Sub
