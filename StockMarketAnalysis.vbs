Sub StockMarketAnalysis()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS AND ADD HEADERs + INFOS
    ' --------------------------------------------
    For Each ws In Worksheets

'Headers

       ws.Cells(1, 9).Value = "Ticker Type"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       ws.Cells(2, 15).Value = "Greatest % Increase"
       ws.Cells(3, 15).Value = "Greatest % Decrease"
       ws.Cells(4, 15).Value = "Greteast Total Volume"
       ws.Cells(1, 16).Value = "Ticker Type"
       ws.Cells(1, 17).Value = "Value"
       
       
 ' Set initial variables for holding the Ticker Type and their Total Stock Value
  
  Dim Ticker_Type As String
  Dim Total_StockVolume As Double
 
  
  'Set initial Value of Stock Volume
  
  Total_StockVolume = 0
  
  
 ' Keep track of the location for each credit card brand in the summary table
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  
  ' Loop through all cells to record Ticker type + Adding their value
  
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
  For i = 2 To lastrow

    ' Check if we are still within the same credit card brand, if we are not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

  ' Set the Ticker Type
      Ticker_Type = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      Total_StockVolume = Total_StockVolume + ws.Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Type

      ' Print the Brand Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_StockVolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Total_StockVolume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Total_StockVolume = Total_StockVolume + ws.Cells(i, 7).Value

    End If

  Next i

       
  Next ws


End Sub

