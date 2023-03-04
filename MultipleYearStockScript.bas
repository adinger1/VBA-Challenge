Attribute VB_Name = "Module1"
Sub MultipleYearStockScript():

Dim wkst As Worksheet

For Each wkst In ThisWorkbook.Worksheets:

    wkst.Cells(1, 9) = "Ticker"
    wkst.Cells(1, 10) = "Yearly Change"
    wkst.Cells(1, 11) = "Percent Change"
    wkst.Cells(1, 12) = "Total Stock"
    
    
    Dim lastrow As Double
    lastrow = wkst.Cells(Rows.Count, 1).End(xlUp).Row
    Dim openingprice As Double
    Dim closingprice As Double
    Dim evaltablerow As Integer
  
    Dim ticker As String
    Dim percentchange As Double
    ticker = wkst.Cells(2, 1)
    openingprice = 0
    evaltablerow = 2
    closingprice = 0
    For i = 2 To lastrow:
        If i = 2 Then
            ticker = wkst.Cells(i, 1)
            openingprice = wkst.Cells(i, 3)
            closingprice = wkst.Cells(i, 6)
            'Adds volume to eval table.
            wkst.Cells(evaltablerow, 12) = wkst.Cells(i, 7)
            
        Else
            'checks if ticker of current row has changed
            If wkst.Cells(i, 1) <> ticker Then
                'fills in table values for that ticker
                wkst.Cells(evaltablerow, 9) = ticker
                wkst.Cells(evaltablerow, 10) = closingprice - openingprice
                
                'Conditional formatting for Yearly Change column
                If wkst.Cells(evaltablerow, 10).Value < 0 Then
                    wkst.Cells(evaltablerow, 10).Interior.ColorIndex = 3
                Else
                    wkst.Cells(evaltablerow, 10).Interior.ColorIndex = 4
                End If
                
                If openingprice <> 0 Then
                    percentchange = (closingprice - openingprice) / openingprice
                Else
                    percentchange = 0
                End If
                
                wkst.Cells(evaltablerow, 11) = FormatPercent(percentchange, 2)
                
                'reset variables for new ticker
                ticker = wkst.Cells(i, 1)
                openingprice = wkst.Cells(i, 3)
                closingprice = wkst.Cells(i, 6)
                evaltablerow = evaltablerow + 1
                wkst.Cells(evaltablerow, 12) = wkst.Cells(i, 7)
                    
            'else statement executes if ticker has NOT changed
            Else
                wkst.Cells(evaltablerow, 12) = wkst.Cells(evaltablerow, 12) + wkst.Cells(i, 7)
                closingprice = wkst.Cells(i, 6)
                If i = lastrow Then
                    wkst.Cells(evaltablerow, 9) = ticker
                    wkst.Cells(evaltablerow, 10) = closingprice - openingprice
                    If wkst.Cells(evaltablerow, 10).Value < 0 Then
                        wkst.Cells(evaltablerow, 10).Interior.ColorIndex = 3
                    Else
                        wkst.Cells(evaltablerow, 10).Interior.ColorIndex = 4
                    End If
                    
                    If openingprice <> 0 Then
                        percentchange = (closingprice - openingprice) / openingprice
                    Else
                        percentchange = 0
                    End If
                    wkst.Cells(evaltablerow, 11) = FormatPercent(percentchange, 2)
                End If
            End If
    
        End If
    Next i
    
    wkst.Range("N2").Value = "Greatest % Increase"
    wkst.Range("N3").Value = "Greatest % Decrease"
    wkst.Range("N4").Value = "Greatest Total Volume"
    wkst.Range("O1").Value = "Ticker"
    wkst.Range("P1").Value = "Value"
    
    
    'Now represents the last row of the summary table
    lastrow = wkst.Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To lastrow:
        If wkst.Cells(i, 11) > greatestincrease Then
            wkst.Range("O2").Value = wkst.Cells(i, 9)
            wkst.Range("P2").Value = FormatPercent(wkst.Cells(i, 11), 2)
        End If
        
        If wkst.Cells(i, 11) < greatestdecrease Then
            wkst.Range("O3").Value = wkst.Cells(i, 9)
            wkst.Range("P3").Value = FormatPercent(wkst.Cells(i, 11), 2)
        End If
        
        If wkst.Cells(i, 12) > greatestvolume Then
            wkst.Range("O4").Value = wkst.Cells(i, 9)
            wkst.Range("P4").Value = wkst.Cells(i, 12)
        End If
    Next i

Next wkst

End Sub


