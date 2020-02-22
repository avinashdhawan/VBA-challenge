Attribute VB_Name = "Module1"
Sub StockMarket():

 'Used to run macros on each worksheet
 For Each ws In Worksheets
 
    'Define Ticker symbol as a variable to summarize stock data ticker symbol
    Dim Ticker As String
    
    'Define Opening Price of the stock as a variable as currency
    Dim Op_Price As Double
    
    'Define Closing Price of the stock as a variable as currency
    Dim Cl_Price As Double
    
    'Define Price Change of the stock as a variable as currency
    Dim Price_Change As Double
    
    'Define Percent change of Price of the stock as a variable as double
    Dim Pct_Change As Double
    
    'Define of the volume of stock traded as a long variable
    Dim Volume1 As Double
    
    'Counter used to view each row of spreadsheet tab
    Dim i As Double
    
    'Store last row of spreadsheet as a variable
    Dim lastRow As Double

    'Use counter to have sequential tickers one after another to summarize in table
    Dim SummaryRow As Double
    
    
    
    'Determine last row in spreadsheet with data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'First row where summarized data goes

        SummaryRow = 2
        
        'Set Ticker symbol volume to zero
        Volume1 = 0
        
       ws.Range("J1").Value = "Ticker"
       ws.Range("K1").Value = "Price Change"
       ws.Range("L1").Value = "Percent Change"
       ws.Range("M1").Value = "Total Volume"
       
       ws.Range("O1").Value = "First Opening Price"
       ws.Range("P1").Value = "Last Closing Price"
        
    'For loop that views all rows from row 2 to the very last row in sheet
    For i = 2 To (lastRow)
    
      'If the next cell does not equal the current cell then do this
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

   ' Ticker symbol defined
      Ticker = ws.Cells(i, 1).Value
      Op_Price = ws.Cells(i + 1, 3).Value
        
      'Ticker closing value at end of year
        
        Cl_Price = ws.Cells(i, 6).Value
        
        'Volume of Ticker
        
         Volume1 = Volume1 + ws.Cells(i, 7).Value
        
    ws.Range("O" & SummaryRow + 1).Value = Op_Price
      ws.Range("O" & 2).Value = ws.Cells(2, 3).Value
      
      

      ws.Range("P" & SummaryRow).Value = Cl_Price

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("J" & SummaryRow).Value = Ticker

      Price_Change = ws.Range("P" & SummaryRow) - ws.Range("O" & SummaryRow)
      ' Print the difference in prices to the Summary Table
      ws.Range("K" & SummaryRow).Value = Price_Change
      
      If ws.Range("O" & SummaryRow) > 0 Then
               Pct_Change = Price_Change / ws.Range("O" & SummaryRow)
       Else
            Pct_Change = 0
      End If
            
        ws.Range("L" & SummaryRow).Value = Pct_Change
    
      If Price_Change < 0 Then
        ws.Range("K" & SummaryRow).Interior.ColorIndex = 3
      ElseIf Price_Change >= 0 Then
        ws.Range("K" & SummaryRow).Interior.ColorIndex = 4
      End If
        
      
    
            'Print the difference in prices to the Summary Table
          
 
        
            ' Print the total volume to the Summary Table
      ws.Range("M" & SummaryRow).Value = Volume1
      
     
     
      
      ' Add one to the summary table row
      SummaryRow = SummaryRow + 1
     
      
      Volume1 = 0
      
 Else
  
     Volume1 = Volume1 + ws.Cells(i, 7).Value
   
     
  End If

  Next i


 Next ws

End Sub
