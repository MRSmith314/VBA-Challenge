'--------------------------------------------------------------------------------------
    ' CREATE A SCRIPT TO LOOP THROUGH ALL STOCKS FOR ONE YEAR AND OUTPUTS THE FOLLOWING:
        ' TICKER SYMBOL
        ' YEARLY CHANGE FROM OPEING PRICE TO CLOSING PRICE FOR EACH YEAR
        ' PERCENTAGE CHANGE FROM THE OPEING PRICE TO THE CLOSING PRICE FOR EACH YEAR
        ' TOTAL STOCK VOLUME FOR EACH TICKER SYMBOL
'---------------------------------------------------------------------------------------
        
Sub StockAnalysis():
    
    ' Define variables to be used for each worksheet
    
    For Each ws In Worksheets
    
    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
     TotalVolume = 0
                
    ' Determine last row
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    ' Create Summary Table for each Ticker symbol
    Dim TickerSummaryTable As Integer
    TickerSummaryTable = 2
              
    Dim OpenPrice As Double
    Dim ClosePrice As Double
                
    ' Create Summary for Max Increase and decrease
  
    Dim MaxPercentChange As Double
    Dim MinPercentChange As Double
    Dim MaxVolume As LongLong
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim MaxVolumeTicker As String
  
    MaxPercentChange = 0
    MinPercentChange = 0
    MaxVolume = 0
    OpenPrice = ws.Cells(2, 3).Value
            
    ' Create Summary Table labels
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
  
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
    ' Loop through all rows and insert values in Summary Table
   
    For i = 2 To LastRow
   
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        TickerSymbol = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ClosePrice = ws.Cells(i, 6).Value
        ws.Range("I" & TickerSummaryTable).Value = TickerSymbol
            
        YearlyChange = ClosePrice - OpenPrice
        PercentChange = (YearlyChange / OpenPrice)
      
        ws.Range("J" & TickerSummaryTable).Value = YearlyChange
        ws.Range("K" & TickerSummaryTable).Value = PercentChange
        ws.Range("L" & TickerSummaryTable).Value = TotalVolume

    ' Format Yearly Change column for positive and negative change
        
     If YearlyChange >= 0 Then
     ws.Range("J" & TickerSummaryTable).Interior.ColorIndex = 4 'Green
            
      Else
      ws.Range("J" & TickerSummaryTable).Interior.ColorIndex = 3 'Red
            
      End If
        
' Calculate Greatest % Changes

      If PercentChange > MaxPercentChange Then
      MaxPercentChange = PercentChange
      MaxTicker = TickerSymbol
            
      ElseIf PercentChange < MinPercentChange Then
      MinPercentChange = PercentChange
      MinTicker = TickerSymbol
                
      End If
        
' Calculate Greatest Total volume

      If TotalVolume > MaxVolume Then
      MaxVolume = TotalVolume
      MaxVolumeTicker = TickerSymbol
            
      End If
        
' Iterate Summary rows

      TickerSummaryTable = TickerSummaryTable + 1
        
' Reset Values

      TotalVolume = 0
      PercentChange = 0
      OpenPrice = ws.Cells(i + 1, 3).Value
            
    Else
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
    End If
            
  Next i
        
' Insert Greatest Increase/Descrease Values in seperate summary

  ws.Range("P2").Value = MaxTicker
  ws.Range("P3").Value = MinTicker
  ws.Range("P4").Value = MaxVolumeTicker
  ws.Range("Q2").Value = MaxPercentChange
  ws.Range("Q3").Value = MinPercentChange
  ws.Range("Q4").Value = MaxVolume
        
' Format numbers to two decimal places with % symbol and AutoFit columns

 ws.Range("Q2").NumberFormat = "0.00%"
 ws.Range("Q3").NumberFormat = "0.00%"
 
 ws.Columns("I:Q").AutoFit
            
Next ws
    
End Sub

