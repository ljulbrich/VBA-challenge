Attribute VB_Name = "Module1"

' Written by Lucas Ulbrich

Sub vba_challenge():

    Dim ws As Worksheet
    Dim tickerVolume As Double
    Dim endOfPage As Long
    Dim n As Integer
    
    Dim yearlyChangeOpen As Double
    Dim yearlyChangeClose As Double
    Dim yearlyChange As Double
    
    Dim percentageChange As Integer
    Dim percentageCounter As Integer
    Dim tickerCompare As Integer

    ' Begin workbook loop.
    
    For Each ws In ThisWorkbook.Worksheets
        n = 2
        endOfPage = (ws.Cells(Rows.Count, 1).End(xlUp).Row)
        
        ws.Cells(1, 9).Value = "Ticker symbol"
        ws.Cells(1, 10).Value = "Yearly change"
        ws.Cells(1, 11).Value = "Percentage change"
        ws.Cells(1, 12).Value = "Total Volume"
        For j = 2 To endOfPage
        
            tickerVolume = tickerVolume + ws.Cells(j, 7).Value
            yearlyChangeOpen = ws.Cells(j, 3).Value
            percentageCounter = percentageCounter + 1
            
            If CStr(ws.Cells(j + 1, 1).Value) <> CStr(ws.Cells(j, 1).Value) Then
                yearlyChangeClose = ws.Cells(j, 6).Value
                yearlyChange = yearlyChangeOpen - yearlyChangeClose
                percentChange = yearlyChange * 100

                
                ' Fill three columns with
                ws.Cells(n, 9).Value = ws.Cells(j, 1).Value
                ws.Cells(n, 10).Value = yearlyChange
                ws.Cells(n, 11).Value = percentChange
                ws.Cells(n, 12).Value = tickerVolume
                
                ' Conditional formatting
                If ws.Cells(n, 10).Value > 0 Then   ' Yearly change
                    ws.Cells(n, 10).Interior.ColorIndex = 4    ' Green for positive change
                    ws.Cells(n, 11).Interior.ColorIndex = 4
                    ws.Cells(n, 12).Interior.ColorIndex = 4
                ElseIf ws.Cells(n, 10).Value < 0 Then
                    ws.Cells(n, 10).Interior.ColorIndex = 3    ' Red for negative change
                    ws.Cells(n, 11).Interior.ColorIndex = 3
                    ws.Cells(n, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(n, 10).Interior.ColorIndex = 6    ' Yellow for no change
                    ws.Cells(n, 11).Interior.ColorIndex = 6
                    ws.Cells(n, 12).Interior.ColorIndex = 6
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
