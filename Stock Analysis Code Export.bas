Attribute VB_Name = "Module1"
 

Sub StockAnalysis():


For Each ws In Worksheets

            Dim Ticker As Integer
                Dim Day As Integer
                Dim Opening As Integer
                Dim Closing As Integer
                Dim Volume As Integer
                Dim TotalStockVolume As Double
                Dim OpeningValue As Double
                Dim ClosingValue As Double
                Dim AnnualVariance As Double
                Dim PercentChange As String
            
                'Naming each column by what they represent
                Ticker = 1
                Day = 2
                Opening = 3
                Closing = 6
                Volume = 7
            
                OpeningValue = ws.Cells(2, 3).Value
                ClosingValue = 0
                AnnualVariance = 0
                PercentChange = 0
                TotalStockVolume = 0
            
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 9).Interior.ColorIndex = 25
                ws.Cells(1, 9).Font.ColorIndex = 2
            
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 10).Interior.ColorIndex = 25
                ws.Cells(1, 10).Font.ColorIndex = 2
            
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 11).Interior.ColorIndex = 25
                ws.Cells(1, 11).Font.ColorIndex = 2
            
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Cells(1, 12).Interior.ColorIndex = 25
                ws.Cells(1, 12).Font.ColorIndex = 2

           ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
           
            
            For t = 2 To ws.Cells(Rows.Count, 2).End(xlUp).Row
                If ws.Cells(t + 1, Ticker).Value = ws.Cells(t, Ticker).Value Then
                    TotalStockVolume = ws.Cells(t, Volume).Value + TotalStockVolume
        
        
                ElseIf ws.Cells(t + 1, Ticker).Value <> ws.Cells(t, Ticker).Value Then
                    ClosingValue = ws.Cells(t, Closing).Value
                    ws.Cells(ws.Cells(Rows.Count, 9).End(xlUp).Row + 1, 9).Value = ws.Cells(t + 1, Ticker).Value
                    
                    AnnualVariance = ClosingValue - OpeningValue
                    ws.Cells(ws.Cells(Rows.Count, 10).End(xlUp).Row + 1, 10).Value = AnnualVariance
                        If AnnualVariance > 0 Then
                           ws.Cells(ws.Cells(Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 4
        
                           ElseIf AnnualVariance < 0 Then
                                ws.Cells(ws.Cells(Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 3
                        End If
        
                    PercentChange = FormatPercent(Round(AnnualVariance / OpeningValue, 4))
                    ws.Cells(ws.Cells(Rows.Count, 11).End(xlUp).Row + 1, 11).Value = PercentChange
        
        
        
                    TotalStockVolume = ws.Cells(t, Volume).Value + TotalStockVolume
                    ws.Cells(ws.Cells(Rows.Count, 12).End(xlUp).Row + 1, 12).Value = TotalStockVolume
        
        
        
                    OpeningValue = ws.Cells(t + 1, Opening).Value
                    ClosingValue = 0
                    AnnualVariance = 0
                    PercentChange = 0
                    TotalStockVolume = 0
        
                End If
        
            Next t
        
        
            
            ws.Cells(2, 15).Value = "Greatest Increase"
            ws.Cells(3, 15).Value = "Greatest Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Value"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 16).Interior.ColorIndex = 25
            ws.Cells(1, 16).Font.ColorIndex = 2
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(1, 17).Interior.ColorIndex = 25
            ws.Cells(1, 17).Font.ColorIndex = 2
            
            
           ws.Cells(2, 17).Value = ws.Cells(2, 11).Value
           ws.Cells(3, 17).Value = ws.Cells(2, 11).Value
           ws.Cells(4, 17).Value = ws.Cells(2, 12).Value
              
            
            For n = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
            
            If ws.Cells(n, 11).Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 17).Value = ws.Cells(n, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(n, 9).Value
            End If
            
            If ws.Cells(n, 11).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = ws.Cells(n, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(n, 9).Value
            End If
            
            If ws.Cells(n, 12).Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 17).Value = ws.Cells(n, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
            End If
            
            Next n
    
Next ws


End Sub

