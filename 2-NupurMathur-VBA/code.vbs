Sub analyst()
' Create a script that will loop through all the stocks and take the following info.
    ' Yearly change from what the stock opened the year at to what the closing price was.
    ' The percent change from the what it opened the year at to what it closed.
    ' The total Volume of the stock
    ' Ticker symbol
    'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
' You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    'define variables
    Dim ws As Worksheet
    Dim tickerCol As Integer
    Dim volCol As Integer
    Dim openCol As Integer
    Dim closeCol As Integer
    Dim lRow As Long
    Dim i As Long

    'Set columns
    tickerCol = 1
    volCol = 7
    openCol = 3
    closeCol = 6

    Dim currentTicker As String
    Dim nextTicker As String
    Dim sumTicker As String 'Ticker value for summary printout
    Dim yrChange As Double 'calculation for difference between starting open to ending close
    Dim prChange As Double 'calculation for change between starting open to ending close

    Dim vol As Variant 'volume per ticker
    Dim totalVol As Variant 'calculating total volume per tickerr
    Dim sumVol As Variant 'highest total volume in all tickers
    Dim volTicker As String 'ticker with highest volume
    
    Dim incTicker As String 'ticker with highest increase
    Dim inc As Variant 'increase value
    
    Dim decTicker As String 'ticker with highest Decrease
    Dim dec As Variant 'decrease value
    
    Dim c As Long 'counter
    Dim startPrice As Double
    Dim endPrice As Double

    'Summarize each Worksheet in Workbook
    For Each ws In Worksheets

        'Set counter and price
        c = 1
        startPrice = ws.Range("C2").Value

        ' Find last Row
        lRow = ws.Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row

        ' Set up summary charts
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Gratest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Find ticker symbols and volume then print summary
        For i = 2 To lRow
            
            currentTicker = ws.Cells(i, tickerCol).Value
            nextTicker = ws.Cells(i + 1, tickerCol).Value
            vol = ws.Cells(i, volCol).Value

            ' If the ticker is the same, add its volume to the total
            If currentTicker = nextTicker Then
                totalVol = totalVol + vol
            
            ' If the ticker is not the same:
            ElseIf currentTicker <> nextTicker Then
                ' add last ticker volume to total
                totalVol = totalVol + vol
                ' find close price
                endPrice = ws.Cells(i, closeCol).Value
                
                ' increase total ticker counter
                c = c + 1

                ' Set variables for sunmary row
                sumTicker = currentTicker
                yrChange = endPrice - startPrice
                ' Conditional if divisor is 0
                If startPrice > 0 Then
                    prChange = (endPrice - startPrice) / startPrice
                Else
                    prChange = 0
                End If

                ' print summary row for ticker
                ws.Cells(c, 10).Value = sumTicker
                ws.Cells(c, 11).Value = yrChange
                ws.Cells(c, 12).Value = prChange
                ws.Cells(c, 13).Value = totalVol
                
                ' Find Greatest Increase and Decrease
                If inc < prChange Then
                    inc = prChange
                    incTicker = sumTicker
                ElseIf dec > prChange Then
                    dec = prChange
                    decTicker = sumTicker
                End If

                ' Find Greatest Volume
                If sumVol < totalVol Then
                    sumVol = totalVol
                    volTicker = sumTicker
                End If

                ' zero volume
                totalVol = 0
                ' reset start price
                startPrice = ws.Cells(i + 1, openCol).Value
            End If
        Next i
        
        ' Insert values into summary
        ws.Cells(2, 16).Value = incTicker
        ws.Cells(2, 17).Value = inc
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 16).Value = decTicker
        ws.Cells(3, 17).Value = dec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = volTicker
        ws.Cells(4, 17).Value = sumVol
        ws.Cells(4, 17).NumberFormat = "0"

        ' Conditional formatting for yearly Change
        For i = 2 To c
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i

        ' Format from currency to Numbers and adjust column width
        ws.Columns("J:Q").AutoFit
        ws.Columns(11).NumberFormat = "@"
        ws.Columns(11).HorizontalAlignment = xlRight
        ws.Columns(12).NumberFormat = "0.00%"
        ws.Columns(13).NumberFormat = "0"

    Next ws

End Sub