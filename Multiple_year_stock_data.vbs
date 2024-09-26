Sub SummarizeWorksheets()

    Dim WorksheetCount As Integer
    Dim I As Integer
    
    WorksheetCount = ActiveWorkbook.Worksheets.Count    'Count how many worksheets are contained in the workbook

    For I = 1 To WorksheetCount                         'Loop to cycle through worksheets
        ActiveWorkbook.Worksheets(I).Activate
        PopulateSummaryTable                            'Function to summarize each worksheet
    Next I
    
    ActiveWorkbook.Worksheets(1).Activate               'Return to the first worksheet

End Sub

Function PopulateSummaryTable()
    
    'Setup summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quaterly change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("I:L").ColumnWidth = 18
        
    Dim LastRecord, NumberOfTickers, TickerIndex, NextTickerStartRow As Long
    Dim Tickers() As Variant
    Dim Ticker As Variant
    
    LastRecord = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    'Find last row with data
    
    Tickers = WorksheetFunction.Unique(Range("A2:A" & LastRecord))      'Populate Tickers array with unique ticker values
    
    NumberOfTickers = UBound(Tickers, 1)                  'Count number of unique Tickers
    Range("I2:I" & u + 1).Value = Tickers   '+1 because starting at row 2 and subtracting 1 for array index
    
    TickerIndex = 0
    NextTickerStartRow = 0
    
    
    'Loop through all records to calculate and populate the summary table
    For I = 1 To LastRecord
        If Cells(I + 2, 1).Value <> Cells(I + 1, 1).Value Then      'End of current ticker records triggers calculations and populating the summary table
            TickerIndex = TickerIndex + 1
            QuarterlyChange = Cells(I + 1, 6).Value - Cells(I - RowCount + 1, 6).Value
            Cells(TickerIndex + 1, 10).Value = QuarterlyChange
            PercentChange = QuarterlyChange / Cells(I - RowCount + 1, 6).Value
            Cells(TickerIndex + 1, 11).Value = PercentChange
            TotalStockVolume = 0
            For j = I - RowCount To I
                TotalStockVolume = TotalStockVolume + Cells(j + 1, 7).Value
            Next j
            Cells(TickerIndex + 1, 12).Value = TotalStockVolume
            RowCount = 0
            NextTickerStartRow = I + 2
        Else
           RowCount = RowCount + 1      'Still on same ticker records so incrementing the row count for the ticker
        End If
    Next I

    'Formatting the summary table
    For x = 1 To NumberOfTickers
        If Cells(x + 1, 10).Value < 0 Then
            Cells(x + 1, 10).Interior.ColorIndex = 3
            Cells(x + 1, 11).Interior.ColorIndex = 3
        Else
            Cells(x + 1, 10).Interior.ColorIndex = 4
            Cells(x + 1, 11).Interior.ColorIndex = 4
        End If
        Cells(x + 1, 10).NumberFormat = "_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* ""-""??_);_(@_)"
        Cells(x + 1, 11).NumberFormat = "0.00%"
    Next x
    
    'Creating and formatting the biggest movers table
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("O:Q").ColumnWidth = 21
    
    Dim IncreasePct, DecreasePct As Double
    Dim Volume As LongLong
    
    IncreasePct = 0
    DecreasePct = 0
    Volume = 0
    
    'Looping through the summary table to identify the biggest movers
    For y = 1 To NumberOfTickers
        If Cells(y + 1, 11).Value < 0 And Cells(y + 1, 11).Value < DecreasePct Then
            DecreaseTicker = Cells(y + 1, 9).Value
            DecreasePct = Cells(y + 1, 11).Value
        End If
        If Cells(y + 1, 11).Value >= 0 And Cells(y + 1, 11).Value > IncreasePct Then
            IncreaseTicker = Cells(y + 1, 9).Value
            IncreasePct = Cells(y + 1, 11).Value
        End If
        If Cells(y + 1, 12).Value > Volume Then
            VolumeTicker = Cells(y + 1, 9).Value
            Volume = Cells(y + 1, 12).Value
        End If
    Next y

    'Populating the biggest movers table
    Range("P2").Value = IncreaseTicker
    Range("Q2").Value = IncreasePct
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").Value = DecreaseTicker
    Range("Q3").Value = DecreasePct
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").Value = VolumeTicker
    Range("Q4").Value = Volume


End Function

