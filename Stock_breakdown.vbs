Sub stock_breakdown()
    Dim w As Worksheet
    For Each w In ActiveWorkbook.Worksheets
    w.Activate
    
    Dim i As Long
    Dim ticker_name As String
    Dim summary_row As Integer
    Dim totalrows As Long
    Dim year_open, year_close, total_volume, percent_change As Double
    
    year_close = 0
    summary_row = 2
    totalrows = Cells(1, 1).End(xlDown).Row
    year_open = Cells(2, 3).Value
    total_volume = 0
    
    For i = 2 To totalrows
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_name = Cells(i, 1).Value
            total_volume = total_volume + Cells(i, 7).Value
            year_close = Cells(i, 6).Value
                        
            Range("I1") = "Ticker"
            Range("I" & summary_row).Value = ticker_name
                                  
            Range("J1") = "Yearly Change"
            If ((year_close - year_open) > 0) Then
                Range("J" & summary_row).Style = "Currency"
                Range("J" & summary_row).Value = (year_close - year_open)
                Range("J" & summary_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_row).Style = "Currency"
                Range("J" & summary_row).Value = (year_close - year_open)
                Range("J" & summary_row).Interior.ColorIndex = 3
            End If
            If year_close <> 0 And year_open <> 0 Then
            Range("K1") = "Percent Change"
            Range("K" & summary_row).NumberFormat = "#.##%"
            Range("K" & summary_row).Value = (year_close / year_open) - 1
            Else
            End If
            Range("L1") = "Total Stock Volume"
            Range("L" & summary_row).Value = total_volume
                      
            summary_row = summary_row + 1
            year_open = Cells(i + 1, 3).Value
            total_volume = 0
        Else
            total_volume = total_volume + Cells(i, 7).Value
        End If
       
    Next i
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("J1:Q1").Columns.AutoFit
        Range("O1:O4").Columns.AutoFit
        Range("Q4").Rows.AutoFit
        Dim rng As Range
        Dim vol_rng As Range
       
        Set rng = Range("K2:K4000")
        increase = Application.WorksheetFunction.Max(rng)
        Range("Q2").NumberFormat = "#.##%"
        Range("Q2").Value = increase
              
        decrease = Application.WorksheetFunction.Min(rng)
        Range("Q3").NumberFormat = "#.##%"
        Range("Q3").Value = decrease
        
        Set vol_rng = Range("L2:L4000")
        greatest_volume = Application.WorksheetFunction.Max(vol_rng)
        Range("Q4").Value = greatest_volume
        Range("Q2:Q4").Columns.AutoFit
        
        Dim x As Long
        Dim tickerRng, changeRng, volRng As Range
        
        Set tickerRng = Range("I2:I4000")
        Set changeRng = Range("K2:K4000")
        Set volRng = Range("L2:L4000")
        
        For x = 2 To 3
            Range("P" & x).Value = Application.WorksheetFunction.Index(tickerRng, Application.WorksheetFunction.Match(Range("Q" & x).Value, changeRng, 0))
         Next x
        Range("P4").Value = Application.WorksheetFunction.Index(tickerRng, Application.WorksheetFunction.Match(Range("Q4").Value, volRng, 0))
        
        
    Next w
    
End Sub

