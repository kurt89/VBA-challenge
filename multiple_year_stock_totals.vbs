Sub multiple_year_stock_totals()

  ' ------ Kurt Dietrich -----------------------
  ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' ------------------------------------------
    For Each ws In Worksheets
  
      ' Set an initial variable for holding the ticker symbol
      Dim Ticker_Symbol As String
  
      ' Set an initial variable for holding the year's opening stock price
      Dim Opening_Price As Double
      Opening_Price = ws.Cells(2, 3).Value
  
      ' Set an initial variable for holding the year end closing stock price
      Dim Closing_Price As Double
  
      ' Set an initial variable for holding the year's stock price change
      Dim Change_Price As Double
  
      ' Set an initial variable for holding the yearly stock price percent change
      Dim PercentChange_Price As Double
  
      ' Set an initial variable for holding the total volume per stock
      Dim Volume_Total As Double
      Volume_Total = 0

      ' Keep track of the location for each stock in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
  
      ' Print the Ticker Symbol summary table column header
      ws.Cells(1, 9).Value = "Ticker"
  
      ' Print the Yearly Stock Price Change summary table column header
      ws.Cells(1, 10).Value = "Yearly Change"
  
      ' Print the Yearly Stock Price Percent Change summary table column header
      ws.Cells(1, 11).Value = "Percent Change"
  
      ' Print the Yearly Total Stock Volume summary table column header
      ws.Cells(1, 12).Value = "Total Stock Volume"
  
      ' Determine the Last Data Row
      LastDataRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
      ' Loop through all "daily stock activity" data rows
      For I = 2 To LastDataRow

        ' Check if we are still within the same stock, if it is not...
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

          ' Set the Brand name
          Ticker_Symbol = ws.Cells(I, 1).Value
      
          ' Set the closing stock price
          Closing_Price = ws.Cells(I, 6).Value
      
          ' Calculate the yearly price change
          Change_Price = Closing_Price - Opening_Price
      
          ' Check if opening price is zero
          If Opening_Price = 0 Then
             PercentageChange_Price = 1
             
             Else
             
             ' Calculate the yearly price percent change
             PercentChange_Price = Change_Price / Opening_Price
             
          End If
          
          ' Add to the volume total
          Volume_Total = Volume_Total + ws.Cells(I, 7).Value

          ' Print the stock ticker symbol in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      
          ' Print the yearly price change in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Change_Price
      
            ' Check for positive or negative price change and format by color
            If Change_Price >= 0 Then
        
              ' Set the Cell Color to Green
              ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
          
            Else
        
              ' Otherwise set the Cell Color to Red
              ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
          
            End If
      
          ' Print the yearly percent price change in the Summary Table
          ws.Range("K" & Summary_Table_Row).Value = PercentChange_Price
          ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
          ' Print the volume total to the Summary Table
          ws.Range("L" & Summary_Table_Row).Value = Volume_Total

          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the opening stock price
          Opening_Price = ws.Cells(I + 1, 3).Value
      
          ' Reset the Price Change
          ' Change_Price = 0
      
          ' Reset the % Price Change
          ' PercentChange_Price = 0
          
          ' Reset the Volume Total
          Volume_Total = 0

        ' If the cell immediately following a row is the same stock...
        Else

          ' Add to the Volume Total
          Volume_Total = Volume_Total + ws.Cells(I, 7).Value

        End If
        
    Next I
    
    ' ---------------------------------------------------------------------------
    ' ------ *** BONUS CODE ***          ----------------------------------------
    ' ------ Return the stock with the:  ----------------------------------------
    ' ------  (1) Greatest % Increase    ----------------------------------------
    ' ------  (2) Greatest & Decrease    ----------------------------------------
    ' ------  (3) Greatest Total Volume  ----------------------------------------
    ' ---------------------------------------------------------------------------
  
        ' Print the Ticker Symbol column header for the Min Max Table
        ws.Cells(1, 16).Value = "Ticker"
  
        ' Print the Value column header for the Min Max Table
        ws.Cells(1, 17).Value = "Value"
   
        ' Print the Greatest Percent Increase row header for the Min Max Table
        ws.Cells(2, 15).Value = "Greatest % Increase"
  
        ' Print the Greatest Percent Decrease row header for the Min Max Table
        ws.Cells(3, 15).Value = "Greatest % Decrease"
  
        ' Print the Greatest Total Volume row header for the Min Max Table
        ws.Cells(4, 15).Value = "Greatest Total Volume"
         
        ' Determine the Last Summary Data Row
        Dim LastSummaryRow As Integer
        LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Define the range of all the last cell in the summary % change column
        ' Dim Range_Last_Pct_Cell As Range
        ' Range_Last_Pct_Cell = ws.Range("K" & LastSummaryRow)
    
        ' Define the range of all Yearly Price % Changes to measure increases
        ' Dim Range_Pct_Inc As Range
        ' Range_Pct_Inc = ws.Range(k2, Range_Last_Pct_Cell)
    
        ' Define the range of all Yearly Price % Changes to measure decreases
        ' Dim Range_Pct_Dec As Range
        ' Range_Pct_Dec = ws.Range(k2, Range_Last_Pct_Cell)
    
        ' Define the range of all Total Stock Volumes
        ' Dim Range_Vol_Sum As Range
        ' Range_Vol_Sum = ws.Range(l2, "L" & LastSummaryRow)
    
        ' Set a variable as the greatest yearly stock price percent increase
        Dim GreatestPctIncrease_Price As Double
    
        ' Determine the Value of the Greatest Stock Price % Increase
        ' GreatestPctIncrease_Price = WorksheetFunction.Max(Range_Pct_Inc.Value)
        GreatestPctIncrease_Price = WorksheetFunction.Max(ws.[k2:k3200].Value)
        ' GreatestPctIncrease_Price = WorksheetFunction.Max(ws.[k2:].Value)
    
        ' Set a variable as the greatest yearly stock price percent decrease
        Dim GreatestPctDecrease_Price As Double
    
        ' Determine the Value of the Greatest Stock Price % Decrease
        ' GreatestPctDecrease_Price = WorksheetFunction.Min(Range_Pct_Dec.Value)
        GreatestPctDecrease_Price = WorksheetFunction.Min(ws.[k2:k3200].Value)
        ' GreatestPctDecrease_Price = WorksheetFunction.Min(ws.[k2:"K" & LastSummaryRow].Value)
    
        ' Set a variable as the greatest annual volume
        Dim Greatest_Volume As Double
    
        ' Determine the Value of the Greatest Total Volume
        ' Greatest_Volume = WorksheetFunction.Max(Range_Vol_Sum.Value)
        Greatest_Volume = WorksheetFunction.Max(ws.[L2:L3200].Value)
        ' Greatest_Volume = WorksheetFunction.Max(ws.[l2:"L" & LastSummaryRow].Value)
    
        ' Loop through all "summary stock activity" data rows
    
        For J = 2 To LastSummaryRow
      
          ' Check the Summary Row for the stock with the Greatest Total Volume
          If ws.Cells(J, 12).Value = Greatest_Volume Then
      
             ' If "Yes" then print the Ticker and Value
             ws.Cells(4, 16).Value = ws.Cells(J, 9).Value
             ws.Cells(4, 17).Value = ws.Cells(J, 12).Value
        
          ' Check the Summary Row for the stock with the Greatest % Increase
          ElseIf ws.Cells(J, 11).Value = GreatestPctIncrease_Price Then
         
             ' If "Yes" then print the Ticker and Value
             ws.Cells(2, 16).Value = ws.Cells(J, 9).Value
             ws.Cells(2, 17).Value = ws.Cells(J, 11).Value
             ws.Cells(2, 17).NumberFormat = "0.00%"
      
          ' Test of the Summary Row is the stock with the Greatest % Decrease
          ElseIf ws.Cells(J, 11).Value = GreatestPctDecrease_Price Then
      
             ' If "Yes" then print the Ticker and Value
             ws.Cells(3, 16).Value = ws.Cells(J, 9).Value
             ws.Cells(3, 17).Value = ws.Cells(J, 11).Value
             ws.Cells(3, 17).NumberFormat = "0.00%"
         
        End If
   
        Next J
        
    Next ws
      
End Sub
  