Sub WolfofWallStreet()

'Create variables for ticker symbol, yearly change, percent change, and total stock volume
Dim tickersym As String
Dim yearChange As Integer
Dim pctChange As Double
Dim totalVol As Double
Dim stockOpen As Double
Dim newTicker As Integer
Dim WS_Count As Integer
Dim W As Integer


'Set values for numerical data
yearChange = 0
pctChange = 0
totalVol = 0
newTicker = 2
WS_Count = ActiveWorkbook.Worksheets.Count

'Loop through all worksheets
For W = 1 To WS_Count

lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    'Name columns for info
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Stock Volume"
        
    'Loop through rows
    For I = 2 To lastRow
    
        'Check to see if ticker symbol changes
        If Cells(I + 1, 1).Value <> Cells(I, 1) Then
            
            'Set Ticker Name
            tickersym = Cells(I, 1).Value
        
            
            'Determine yearly change
            
            
            'Determine the percentage change
            '(closing percentage change - opening percentage change)/ opening percentage change
            
            
            
            'Determine total stock volume
            totalVol = totalVol + Cells(I, 7).Value
            
            'Display desired info in columns
            Cells(newTicker, 9).Value = tickersym
            Cells(newTicker, 10).Value = yearChange
            Cells(newTicker, 11).Value = pctChange
            Cells(newTicker, 12).Value = totalVol
            
            'Every time ticker changes start new row
            newTicker = newTicker + 1
            
            'Reset total volume
            totalVol = 0
               
        Else
            If Cells(I, 1).Value <> Cells(I - 1, 1).Value Then
            
            'Get value of first open price
            stockOpen = Cells(I, 3).Value
            
            End If
            
            'Total up all stock volume for year
            totalVol = totalVol + Cells(I, 7).Value
            
            
        End If
    
    Next I

Next W

End Sub
