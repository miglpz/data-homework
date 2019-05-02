Sub alphaTesting()

Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim yearlyChange As Double
        Dim ticker As String
        Dim percentChange As Double
        Dim volume As Double
        Dim Summary As Integer
        
        Summary = 2
        
        For i = 2 To lastRow
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                 ticker = Cells(i, 1).Value
                 volume = Cells(i, 7).Value
                 yearOpen = Cells(i, 3).Value
                 yearClose = Cells(i, 6).Value
                 
                 yearlyChange = yearClose - yearOpen
                 percentChange = (yearClose - yearOpen) / yearClose
                 
                Cells(Summary, 9).Value = ticker
                Cells(Summary, 10).Value = yearlyChange
                Cells(Summary, 11).Value = percentChange
                Cells(Summary, 12).Value = volume
            Summary = Summary + 1

             vol = 0
                 
            End If
        Next i
        
    Next ws
    
End Sub
