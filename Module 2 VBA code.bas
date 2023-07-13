<<<<<<< HEAD
Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Declare variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    ' Set initial values
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    ' Initialize variables
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    
    ' Set headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Loop through all the stocks
    For i = 2 To lastRow
        
        ' Check if the ticker has changed
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set ticker symbol
            ticker = Cells(i, 1).Value
            
            ' Set closing price
            closingPrice = Cells(i, 6).Value
            
            ' Set total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            
            ' Calculate yearly change
            openingPrice = Cells(j, 3).Value
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate percent change
            If openingPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openingPrice
            End If
            
            ' Output results
            Range("I" & j).Value = ticker
            Range("J" & j).Value = yearlyChange
            Range("K" & j).Value = percentChange
            Range("L" & j).Value = totalVolume
            Range("O1").Value = "Ticker"
            Range("P1").Value = "Value"
            Range("N1").Value = "Type"
            Range("N2").Value = "Greatest % Increase"
            Range("N3").Value = "Greatest % Decrease"
            Range("N4").Value = "Greatest Total Volume"
            
            ' Update maxIncrease, maxDecrease, and maxVolume
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            Range("O2").Value = ticker
            Range("P2").Value = maxIncrease
        End If
        
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            Range("O3").Value = ticker
            Range("P3").Value = maxDecrease
        End If
        
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            Range("O4").Value = ticker
            Range("P4").Value = maxVolume
        End If
        
            ' Reset variables for next ticker symbol
            j = j + 1
            totalVolume = 0
            
        Else
            
            ' Add to total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub
    

=======
Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Declare variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    ' Set initial values
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    ' Initialize variables
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    
    ' Set headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Loop through all the stocks
    For i = 2 To lastRow
        
        ' Check if the ticker has changed
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set ticker symbol
            ticker = Cells(i, 1).Value
            
            ' Set closing price
            closingPrice = Cells(i, 6).Value
            
            ' Set total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            
            ' Calculate yearly change
            openingPrice = Cells(j, 3).Value
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate percent change
            If openingPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openingPrice
            End If
            
            ' Output results
            Range("I" & j).Value = ticker
            Range("J" & j).Value = yearlyChange
            Range("K" & j).Value = percentChange
            Range("L" & j).Value = totalVolume
            Range("O1").Value = "Ticker"
            Range("P1").Value = "Value"
            Range("N1").Value = "Type"
            Range("N2").Value = "Greatest % Increase"
            Range("N3").Value = "Greatest % Decrease"
            Range("N4").Value = "Greatest Total Volume"
            
            ' Update maxIncrease, maxDecrease, and maxVolume
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            Range("O2").Value = ticker
            Range("P2").Value = maxIncrease
        End If
        
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            Range("O3").Value = ticker
            Range("P3").Value = maxDecrease
        End If
        
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            Range("O4").Value = ticker
            Range("P4").Value = maxVolume
        End If
        
            ' Reset variables for next ticker symbol
            j = j + 1
            totalVolume = 0
            
        Else
            
            ' Add to total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub
    

>>>>>>> 34ba37b4fdafb44690f8d44ae132b1c554a42b35
