Attribute VB_Name = "Module1"
Sub ticker_info():

    Dim ws As Worksheet
    Dim tickerVolume As Double
    Dim endOfPage As Integer
    Dim n As Integer
    
    Dim yearlyChangeOpen As Double
    Dim yearlyChangeClose As Double
    Dim yearlyChange As Double
    
    Dim percentageChange As Integer
    Dim percentageCounter As Integer
    Dim tickerCompare As Integer
    
    n = 2

    ' Begin workbook loop.
    
    For Each ws In ThisWorkbook.Worksheets
        
        endOfPage = (Cells(Rows.Count, 1).End(xlUp).Row)
        
        Cells(1, 9).Value = "Ticker symbol"
        Cells(1, 10).Value = "Yearly change"
        Cells(1, 11).Value = "Percentage change"
        Cells(1, 12).Value = "Total Volume"
        For j = 2 To endOfPage
        
            tickerVolume = tickerVolume + Cells(j, 7).Value
            yearlyChangeOpen = Cells(j, 3).Value
            percentageCounter = percentageCounter + 1
            
            If CStr(Cells(j + 1, 1).Value) <> CStr(Cells(j, 1).Value) Then
                yearlyChangeClose = Cells(j, 6).Value
                yearlyChange = yearlyChangeOpen - yearlyChangeClose
                percentChange = yearlyChange * 100

                
                ' Fill three columns with
                Cells(n, 9).Value = Cells(j, 1).Value
                Cells(n, 10).Value = yearlyChange
                Cells(n, 11).Value = percentChange
                Cells(n, 12).Value = tickerVolume
                
                ' Conditional formatting
                If Cells(n, 10).Value > 0 Then   ' Yearly change
                    Cells(n, 10).Interior.ColorIndex = 4    ' Green for positive change
                    Cells(n, 11).Interior.ColorIndex = 4
                    Cells(n, 12).Interior.ColorIndex = 4
                ElseIf Cells(n, 10).Value < 0 Then
                    Cells(n, 10).Interior.ColorIndex = 3    ' Red for negative change
                    Cells(n, 11).Interior.ColorIndex = 3
                    Cells(n, 12).Interior.ColorIndex = 3
                Else
                    Cells(n, 10).Interior.ColorIndex = 6    ' Yellow for no change
                    Cells(n, 11).Interior.ColorIndex = 6
                    Cells(n, 12).Interior.ColorIndex = 6
                End If
                
                
                ' (Re)set counters
                tickerVolume = 0
                percentageCounter = 0
                n = n + 1
            Else
            End If
            
        Next j
        
    Next ws ' End workbook loop
    
End Sub
