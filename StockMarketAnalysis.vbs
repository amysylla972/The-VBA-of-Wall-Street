
Sub Market()
 Dim Ws As Worksheet
 'For hard solution
 Dim COMMAND_SPREADSHEET As Boolean
  
    COMMAND_SPREADSHEET = True
    
    For Each Ws In Worksheets

        Dim TickerName As String
        Dim TotalVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Percentage As Double
        Dim Summary_Table_Row As Long
        Dim MaxTicker As String
        Dim MinTicker As String
        Dim MaxPercent As Double
        Dim MinPercent As Double
        Dim MaxVolumeTicker As String
        Dim MaxVolume As Double
        Dim Lastrow As Long
        Dim i As Long
          
          'Set Variables
        Percentage = 0
        YearlyChange = 0
        ClosePrice = 0
        TotalVolume = 0
        OpenPrice = 0
        Summary_Table_Row = 2
        MaxTicker = " "
        MinTicker = " "
        MaxPercent = 0
        MaxVolumeTicker = " "
        MaxVolume = 0
        MinPercent = 0
           
        Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
           
           
        'Header for everysheets
  

            
            Ws.Range("I1").Value = "Ticker"
            Ws.Range("J1").Value = "Yearly Change"
            Ws.Range("K1").Value = "Percent Change"
            Ws.Range("L1").Value = "Total Stock Volume"
            Ws.Range("O2").Value = "Greatest % Increase"
            Ws.Range("O3").Value = "Greatest % Decrease"
            Ws.Range("O4").Value = "Greatest Total Volume"
            Ws.Range("P1").Value = "Ticker"
            Ws.Range("Q1").Value = "Value"
     
       'Start of OpenPrice before looping
        OpenPrice = Ws.Cells(2, 3).Value
        
    
        
        ' Loop
        For i = 2 To Lastrow
        
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
                
                TickerName = Ws.Cells(i, 1).Value
                ClosePrice = Ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ' Check Division by 0 condition
                If OpenPrice <> 0 Then
                    Percentage = (YearlyChange / OpenPrice) * 100
                Else
                    ' Unlikely, but it needs to be checked to avoid program crushing
                    MsgBox ("For " & TickerName & ", Row " & CStr(i) & ": Open Price =" & OpenPrice & ". Can't be divided by 0! Check your Open prices")
                End If
                
                TotalVolume = TotalVolume + Ws.Cells(i, 7).Value
              
                
                ' Print the Ticker Name And change
                
                Ws.Range("I" & Summary_Table_Row).Value = TickerName
                Ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                
                ' Change Color
                If (YearlyChange > 0) Then
                    
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                    
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
        
                Ws.Range("K" & Summary_Table_Row).Value = (CStr(Percentage) & "%")
          
                Ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                
        
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset
                YearlyChange = 0
                ClosePrice = 0
                OpenPrice = Ws.Cells(i + 1, 3).Value
              
                
                If (Percentage > MaxPercent) Then
                    MaxPercent = Percentage
                    MaxTicker = TickerName
                    
                ElseIf (Percentage < MinPercent) Then
                    MinPercent = Percentage
                    MinTicker = TickerName
                End If
                       
                If (TotalVolume > MaxVolume) Then
                    MaxVolume = TotalVolume
                    MaxVolumeTicker = TickerName
                End If
                
              
                Percentage = 0
                TotalVolume = 0
                
            
          
            Else
             
                TotalVolume = TotalVolume + Ws.Cells(i, 7).Value
            End If
            
        Next i

            If Not COMMAND_SPREADSHEET Then
            
                Ws.Range("Q2").Value = (CStr(MaxPercent) & "%")
                Ws.Range("Q3").Value = (CStr(MinPercent) & "%")
                Ws.Range("P2").Value = MaxTicker
                Ws.Range("P3").Value = MinTicker
                Ws.Range("Q4").Value = MaxVolume
                Ws.Range("P4").Value = MaxVolumeTicker
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next Ws
End Sub
