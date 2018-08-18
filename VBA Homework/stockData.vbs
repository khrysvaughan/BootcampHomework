Sub stockDataTotal()

'Variables for Easy Solution
Dim ticker As String
Dim nextticker As String
Dim i As Long
Dim endofrow As Long
Dim totalstock As Double
Dim totalmarker As Integer

'Variables for Moderate Solution
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyDiff As Double
Dim percentChange As Double
Dim firstPrice As Boolean

'Variables for Hard Solution
Dim greatIncrease As Double
Dim greatIncreaseStock As String
Dim greatDecrease As Double
Dim greatDecreaseStock As String
Dim greatTotVolume As Double
Dim greatTotVolumeStock As String

For Each ws In Worksheets
    'Easy Solution Initialization
    endofrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    totalstock = 0
    totalmarker = 2
    ticker = ""
    nextticker = ""
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Moderate Solution Iniitialization
    firstPrice = True
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    
    'Hard Solution Initialization
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    greatIncrease = 0
    greatIncreaseStock = ""
    greatDecrease = 0
    greatDecreaseStock = ""
    greatTotVolume = 0
    greatTotVolumeStock = ""
       
    For i = 2 To endofrow
        totalstock = totalstock + ws.Range("G" & i).Value
        ticker = ws.Range("A" & i).Value
        nextticker = ws.Range("A" & i + 1).Value
        
        'Moderate Solution Check
        'If this is the first price of the stock, store it
        If firstPrice = True Then
            openingPrice = ws.Range("C" & i).Value
            firstPrice = False
        End If
        
        If ticker <> nextticker Then
            ws.Range("I" & totalmarker).Value = ticker
            ws.Range("L" & totalmarker).Value = totalstock
                        
            'Moderate Solution
            'Yearly Change
            closingPrice = ws.Range("F" & i).Value
            yearlyDiff = closingPrice - openingPrice
            
            'Percentage of Change
            If openingPrice = 0 Then
                percentChange = 0
            Else
                percentChange = ((closingPrice - openingPrice) / openingPrice)
            End If
            ws.Range("J" & totalmarker).Value = yearlyDiff
            'ws.Range("J" & totalmarker).NumberFormat = "0.000000000"
            ws.Range("K" & totalmarker).Value = percentChange
            ws.Range("K" & totalmarker).NumberFormat = "0.00%"
            
            If yearlyDiff >= 0 Then
                ws.Range("J" & totalmarker).Interior.ColorIndex = 4
            Else
                ws.Range("J" & totalmarker).Interior.ColorIndex = 3
            End If
            
            'Hard Solution
            If percentChange > greatIncrease Then
                greatIncrease = percentChange
                greatIncreaseStock = ticker
            End If
            
            If percentChange < greatDecrease Then
                greatDecrease = percentChange
                greatDecreaseStock = ticker
            End If
            
            If totalstock > greatTotVolume Then
                greatTotVolumeStock = ticker
                greatTotVolume = totalstock
            End If
            
            
            'Reset and Update Variables
            firstPrice = True
            totalmarker = totalmarker + 1
            totalstock = 0
        End If
    Next i
    
    ws.Range("P2").Value = greatIncreaseStock
    ws.Range("Q2").Value = greatIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P3").Value = greatDecreaseStock
    ws.Range("Q3").Value = greatDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = greatTotVolumeStock
    ws.Range("Q4").Value = greatTotVolume
    
    'Autofit the columns
    ws.Columns("A:Q").AutoFit
Next ws
End Sub

