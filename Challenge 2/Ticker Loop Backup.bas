Attribute VB_Name = "Module1"
Sub Ticker_Loop()

Dim ws As Worksheet
Dim Ticker As String
Dim vol As Double
vol = 0
Dim year_open As Double
year_open = 0
Dim year_close As Double
year_close = 0
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
percent_change = 0
Dim Summary_Table_Row As Double



For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    Summary_Table_Row = 2

        For i = 2 To ws.UsedRange.Rows.Count
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_Row, 9).Value = Ticker
                Summary_Table_Row = Summary_Table_Row + 1
    
        End If
        
        Next i

        
Dim ticker_count As Double
ticker_count = 0
Dim ticker_first As Double
ticker_first = 0
Dim ticker_last As Double
ticker_last = 0
        
        ticker_count = ws.Cells(Rows.Count, "I").End(xlUp).Row
        For J = 2 To ticker_count
            
    
            ticker_first = ws.Range("A:A").Find(What:=ws.Cells(J, 9).Value, LookAt:=xlWhole).Row
            ticker_last = ws.Range("A:A").Find(What:=ws.Cells(J, 9).Value, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
    
            year_open = ws.Cells(ticker_first, 3).Value
            year_close = ws.Cells(ticker_last, 6).Value
            vol = Application.WorksheetFunction.Sum(ws.Range("G" & ticker_first & ":G" & ticker_last))
            
            If year_open = 0 Then
            yearly_change = 0
            percent_change = 0
            Else:
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_open
            End If
        
            ws.Cells(J, 10).Value = yearly_change
            ws.Cells(J, 11).Value = percent_change
            ws.Cells(J, 11).Style = "Percent"
            ws.Cells(J, 11).NumberFormat = "0.00%"
            ws.Cells(J, 12).Value = vol
            
            vol = 0
        
            If ws.Cells(J, 10).Value > 0 Then
                ws.Cells(J, 10).Interior.Color = vbGreen
                
            Else
                ws.Cells(J, 10).Interior.Color = vbRed
                
        End If
        
        Next J
    
Dim greatest_increase As Double
greatest_increase = 0
Dim greatest_decrease As Double
greatest_decrease = 0
Dim greatest_volume As Double
greatest_volume = 0

        For k = 2 To ticker_count
        
        
            If ws.Cells(k, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(k, 11).Value
                ws.Cells(2, 17).Value = greatest_increase
                ws.Cells(2, 17).Style = "Percent"
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            End If
        
            Next k
        
        For l = 2 To ticker_count
       
            
            If ws.Cells(l, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(l, 11).Value
                ws.Cells(3, 17).Value = greatest_decrease
                ws.Cells(3, 17).Style = "Percent"
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ws.Cells(3, 16).Value = ws.Cells(l, 9).Value
            End If
            
           Next l
        
        For m = 2 To ticker_count
     
            
            If ws.Cells(m, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(m, 12).Value
                ws.Cells(4, 17).Value = greatest_volume
                ws.Cells(4, 16).Value = ws.Cells(m, 9).Value
            End If
          
            Next m
            
ws.Columns("A:Q").AutoFit

Next ws

End Sub
