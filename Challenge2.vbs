Attribute VB_Name = "Module1"
Sub Stock_Test()


Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
    

    Dim Ticker_Name As String
    Dim Close_Total As Double
    Dim Open_Total As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    Dim Ticker_Summary_Row As Double
    Ticker_Summary_Row = 2
    
    Close_Total = 0
    Open_Total = ws.Cells(2, 3).Value
    Total_Volume = 0
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To RowCount
        
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker_Name = ws.Cells(i, 1).Value
                
                Close_Total = ws.Cells(i, 6).Value
                
                
                
                Percent_Change = ((Close_Total - Open_Total) / Open_Total)
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & Ticker_Summary_Row).Value = Ticker_Name
                
                ws.Range("J" & Ticker_Summary_Row).Value = Close_Total - Open_Total
                
                ws.Range("K" & Ticker_Summary_Row).Value = Percent_Change
                ws.Range("K" & Ticker_Summary_Row).NumberFormat = "0.00%"
                
                ws.Range("L" & Ticker_Summary_Row).Value = Total_Volume
                
                If ws.Range("J" & Ticker_Summary_Row).Value > 0 Then
                    ws.Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 4
                    
                    
                ElseIf ws.Range("J" & Ticker_Summary_Row).Value < 0 Then
                    ws.Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 3
                End If
                
                
                Ticker_Summary_Row = Ticker_Summary_Row + 1
                
                Total_Volume = 0
                
                Open_Total = ws.Cells(i + 1, 3).Value
                
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                ' Close_Total = Close_Total + ws.Cells(i, 6).Value
                
               ' Open_Total = Open_Total + ws.Cells(i, 3).Value
                
                ' Percent_Change = ((Close_Total - Open_Total) / Open_Total) * 100
                
                
            End If
            
        Next i
    
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
    
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
    
    ws.Range("Q3").NumberFormat = "0.00%"
        
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
    
    Max_Index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & RowCount), 0)
    
    Min_Index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & RowCount), 0)
    
    Max_Index_Vol = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & RowCount), 0)
    
    ws.Range("P2").Value = ws.Cells(Max_Index + 1, 9).Value
    
    ws.Range("P3").Value = ws.Cells(Min_Index + 1, 9).Value
    
    ws.Range("P4").Value = ws.Cells(Max_Index_Vol + 1, 9).Value
    
    Next ws
        
End Sub
