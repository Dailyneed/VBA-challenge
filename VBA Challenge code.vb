'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker
'Yearly change from the opening price at the begining of agiven year to the closing price at end of that year
'The percentage change from the opening price at the begining of a given year to the closing price at the end of that year
'The total stock volume
'Use conditional formatting that will highlight positive change in green and negative change in red

Sub Stock_yearlychange()
Dim ws As Worksheet
For Each ws In Worksheets
    'Declare all varibles
    Dim Ticker_symbol, MaxVolume_Ticker, Maxpercent_Ticker, Minpercent_Ticker As String
    Dim Year_Openprice, Year_Closeprice, Percent_change, Yearly_change, Total_Stockvolume As Double
    Dim i, k, Counter, lastrow As Long
    Dim Red, Green As Variant
    Dim Max_totalvolume, Maxpercent, Minpercent, Y As Double
    
    
    'Set variables initial value
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Counter = 2
    k = 2
    Total_Stockvolume = 0
    Max_totalvolume = 0
    Maxpercent = 0
    Minpercent = 0
    
    
    
    For i = 2 To lastrow
    
    'Print ticker symbol in column K and intialize year open price
    Ticker_symbol = ws.Cells(i, 1).Value
    ws.Cells(Counter, 11).Value = Ticker_symbol
    Year_Openprice = ws.Cells(k, 3).Value
    
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
            'Calculate Total_Stockvolume and print value to column N
            Total_Stockvolume = Total_Stockvolume + ws.Cells(i, 7).Value
        Else
            k = i
            Year_Closeprice = ws.Cells(k, 6).Value
            
            'Calculate yearly change and print value to column L and highlight cell
            Yearly_change = Year_Closeprice - Year_Openprice
            ws.Cells(Counter, 12).Value = Yearly_change
            
                If Yearly_change > 0 Then
                  ws.Cells(Counter, 12).Interior.Color = vbGreen
                ElseIf Yearly_change < 0 Then
                  ws.Cells(Counter, 12).Interior.Color = vbRed
                End If
            'Calculate percent change and print value to column M
            Percent_change = Yearly_change / Year_Openprice
            ws.Cells(Counter, 13).Value = Format(Percent_change, "#.##%")
            
            If Percent_change > 0 Then
                ws.Cells(Counter, 13).Interior.Color = vbGreen
            ElseIf Percent_change < 0 Then
                ws.Cells(Counter, 13).Interior.Color = vbRed
            End If
            
            'Find greatest total volume, greatest % increase and greatest % decrease
                Y = Yearly_change / Year_Openprice
                If ws.Cells(Counter, 14).Value > Max_totalvolume Then
                Max_totalvolume = ws.Cells(Counter, 14).Value
                ws.Cells(4, 18).Value = Max_totalvolume
                MaxVolume_Ticker = ws.Cells(Counter, 11).Value
                ws.Cells(4, 17).Value = MaxVolume_Ticker
                End If
                If Y > Maxpercent Then
                Maxpercent = Y
                ws.Cells(2, 18).Value = Format(Maxpercent, "#.##%")
                Maxpercent_Ticker = ws.Cells(Counter, 11).Value
                ws.Cells(2, 17).Value = Maxpercent_Ticker
                End If
                If Y < Minpercent Then
                Minpercent = Y
                ws.Cells(3, 18).Value = Format(Minpercent, "#.##%")
                Minpercent_Ticker = ws.Cells(Counter, 11).Value
                ws.Cells(3, 17).Value = Minpercent_Ticker
                End If
            'Reset variables
            Counter = Counter + 1
            Total_Stockvolume = 0
            k = k + 1
            
         End If
               
    Next i
    
    
    
    'Fill headers and save values
    'Autofit to display data
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Columns("A:R").AutoFit
    
    
Next ws


   
End Sub




