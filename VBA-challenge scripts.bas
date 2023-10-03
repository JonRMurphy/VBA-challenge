Attribute VB_Name = "Module1"
Sub stocks()

'Looping through worksheets
Dim ws As Excel.Worksheet
    For Each ws In Worksheets
        ws.Activate

'Setting ticker to start at row 2
Dim ticker As Long
ticker = 2

'Setting opening and closing price variables
Dim oprice, cprice As Double
oprice = Cells(2, 3).Value
cprice = 0

'Setting up Yearly Change and Percent Change Variables
Dim yearly As Double
Dim percent As Double

'Setting variable for total volume
Dim total As LongLong
total = 0

'Setting up last row variable
Dim lastRow As Long
lastRow = Range("A1").End(xlDown).Row

'For loop including Ticker, Yearly Change, Percent Change, and Total Stock Volume
For i = 2 To lastRow
    
    'If the ticker is the same
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        
        'Adds up volumes to total
        total = total + Cells(i, 7).Value
    
    'If the ticker changes
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Adds last volume to total
        total = total + Cells(i, 7).Value
        
        'Prints total volume
        Cells(ticker, 12).Value = total
        
        'Prints ticker
        Cells(ticker, 9).Value = Cells(i, 1).Value
        
        'Sets closing price
        cprice = Cells(i, 6).Value
        
        'Prints Yearly Change
        yearly = cprice - oprice
        Cells(ticker, 10).Value = yearly
        
        'Prints Percent Change
        percent = yearly / oprice
        Cells(ticker, 11).Value = percent
        
        'Changes opening price to new ticker's opening price
        oprice = Cells(i + 1, 3).Value
        
        'Sets ticker value to next row
        ticker = ticker + 1
        
        'Resets total for the new ticker
        total = 0
        
    End If
    

Next i

Next

End Sub

Sub greatest()

Dim ws As Excel.Worksheet
    For Each ws In Worksheets
        ws.Activate
        
    
'Setting variables for Greatest% increases and decreases
Dim gpi, gpd As Double

'Setting variable for Greatest total volume
Dim gtotal As LongLong

'Setting variables for tickers of gpi, gpd, and gtotal
Dim gpiTicker, gpdTicker, gtotalTicker As String

'Setting up variables to prepare for testing
gpi = Cells(2, 11).Value
gpd = Cells(2, 11).Value
gtotal = Cells(2, 12).Value
gpiTicker = Cells(2, 9).Value
gpdTicker = Cells(2, 9).Value
gtotalTicker = Cells(2, 9).Value
    
'Setting up last row variable for summarized data specifically
Dim lastRow As Long
lastRow = Range("I1").End(xlDown).Row

'For loop for Greatest% increase and decrease, Greatest total volume
For i = 2 To lastRow
    
    For j = 2 To lastRow
        If Cells(i + 1, 12).Value > gtotal Then
            gtotal = Cells(i + 1, 12).Value
            gtotalTicker = Cells(i + 1, 9).Value
        End If
    Next j
    If Cells(i + 1, 11).Value > gpi Then
        gpi = Cells(i + 1, 11).Value
        gpiTicker = Cells(i + 1, 9).Value
    
    ElseIf Cells(i + 1, 11).Value < gpd Then
        gpd = Cells(i + 1, 11).Value
        gpdTicker = Cells(i + 1, 9).Value
    
    End If
    
Next i

'Printing tickers for gpi, gpd, and gtotal
Range("O2").Value = gpiTicker
Range("O3").Value = gpdTicker
Range("O4").Value = gtotalTicker

'Printing Greatest% increase and decrease, Greatest total volume
Range("P2").Value = gpi
Range("P3").Value = gpd
Range("P4").Value = gtotal

Next
 
End Sub


