Sub stockmarket()

sheetValue = InputBox("Enter year of sheet of data you would like to see.")
Worksheets(sheetValue).Activate

'define all variables

Dim ticker As String

Dim open_price As Double
open_price = Cells(2, 3)

Dim close_price As Double

Dim yearly_change As Double

Dim percent_change As Double

Dim stock_volume As Double

'define math

yearly_change = close_price - open_price

percent_change = yearly_change / open_price

'creates table
Dim summary_table_row As Long
summary_table_row = 2

'Inserting Data Via Ranges
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly_Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest%Increase"
    Range("N3").Value = "Greatest%Decrease"
    Range("N4").Value = "Greates Total Volume"
    
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'start loop
For i = 2 To lastrow
    If Cells(i + 1, 1) <> Cells(i, 1).Value Then
    'defining ticker name
        ticker = Cells(i, 1)
        
    'defining stock volume
        stock_volume = stock_volume + Cells(i, 7)
        
    'pulling ticker name to table
    Range("I" & summary_table_row) = ticker
    
    'pulling stock volume data to data
    Range("L" & summary_table_row) = stock_volume
    
    'defining closing price,where it is in the data
    close_price = Cells(i, 6)
    
    'calculate yearly change
    yearly_change = (close_price - open_price)
    
    'put the yearly_change in the table
    Range("J" & summary_table_row) = yearly_change
    
    'percent change
    If (open_price = 0) Then
    
        percent_change = 0
        
    Else
    
        percent_change = yearly_change / open_price
        
    End If
    
    Range("K" & summary_table_row) = percent_change
    Range("K" & summary_table_row).NumberFormat = "0.00%"
    
    summary_table_row = summary_table_row + 1
    
    'reset volume counter
    stock_volume = 0
    
    'reset opening price
    open_price = Cells(i + 1, 3)
    
    Else

     stock_volume = stock_volume + Cells(i, 7)
    
    End If
    
   Next i
   
   'defines new table
lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow_summary_table

    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 10
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i

    For i = 2 To lastrow_summary_table
        
            'Find the maximum percent change
            'pull max from percent change
            
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
            
            'input pull from columns
                Cells(2, 15).Value = Cells(i, 9).Value
                Cells(2, 16).Value = Cells(i, 11).Value
                Cells(2, 16).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 15).Value = Cells(i, 9).Value
                Cells(3, 16).Value = Cells(i, 11).Value
                Cells(3, 16).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 15).Value = Cells(i, 9).Value
                Cells(4, 16).Value = Cells(i, 12).Value
                    
        End If
    Next i



End Sub
