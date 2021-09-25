Attribute VB_Name = "Module1"
Sub Stocks():
    'Define the ticker value in a Variable
     Dim TickerLetter As String
     'Variables for our opening and closing to use for changes
     Dim YearlyOpen As Double
     Dim YearlyClose As Double
     'Variable to help defineChanges and their start
     Dim YOpenStart As LongLong
     'Variable for total stock
     Dim TotalStock As LongLong
     'SummaryTable row counter
     Dim SummaryTableRow As LongLong
    
     For Each ws In Worksheets
      'Set the column Headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        YOpenStart = 2
        TotalStock = 0
        SummaryTableRow = 2
    'Variable for LastRow
          Dim LastRow As LongLong
    'Count the number of Rows
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    'Loop through all rows
            For Row = 2 To LastRow
          ' define yearly open before if statement
            YearlyOpen = ws.Range("C" & YOpenStart).Value
    'Set up ticker name changeswhen they are not equal to one another
            If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) Then
            'Reset
            TickerLetter = ws.Range("A" & Row).Value
            YearlyClose = ws.Range("F" & Row).Value
            
            'Add up the Stock Total Volume
             TotalStock = TotalStock + ws.Range("G" & Row).Value
            
            
            
            
            
            
            'Populate the summary Table
            ws.Range("I" & SummaryTableRow).Value = TickerLetter
            ws.Range("J" & SummaryTableRow).Value = YearlyClose - YearlyOpen
            'seperate If for division because we ran into divide by 0 issues lol
            If YearlyOpen = 0 Then
            ws.Range("K" & SummaryTableRow).Value = Null
            ws.Range("J" & SummaryTableRow).Value = Null
            Else
            'Populate the values once divide by 0 is off the table
            ws.Range("J" & SummaryTableRow).Value = YearlyClose - YearlyOpen
            ws.Range("K" & SummaryTableRow).Value = (YearlyClose - YearlyOpen) / YearlyOpen
            End If
            ws.Range("L" & SummaryTableRow).Value = TotalStock
            'Add one to the summary table count
            SummaryTableRow = SummaryTableRow + 1
            'Reset the Stock Total to 0
            YOpenStart = Row + 1
            TotalStock = 0
           'If the ticker letter is the same compile and sum the total value
            Else
             TotalStock = TotalStock + ws.Range("G" & Row).Value
             
             End If
            'Made a mistake which needed correcting back to original formatting
            ws.Range("L" & SummaryTableRow).NumberFormat = "0"
            'Format the percent change column to be in percentage
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

            'Color Formatting Positives in Green Negatives in Red (Only for yearly change ((Note in README.)
            If ws.Range("J" & SummaryTableRow).Value > 0 Then
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            Else
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
         
         Next Row
             ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
     Next ws
     
End Sub
