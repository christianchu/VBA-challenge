Sub test1()

'Dim ws As Worksheet
Dim Ticker As String
Dim Total_Stock_Volume As Double
'initialize counter'
Total_Stock_Volume = 0

Dim close_date As Double
Dim open_date As Double
open_date = 0
close_date = 0

Dim close_price As Integer
Dim open_price As Integer
open_price = 0
close_price = 0

'Low = Application.WorksheetFunction.Min(Worksheets)
'High = Application.WorksheetFunction.Max(Worksheets)

Dim New_Table As Integer
New_Table = 2

'For Each ws In Worksheets

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'create headers'
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

    For i = 2 To LastRow
        'conditional if ticker changes'
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'store ticker data'
        Ticker = Cells(i, 1).Value
        'store total stock vol'
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        'store open & close date value'
        open_date = WorksheetFunction.Min(Range("B:B").Value)
        close_date = WorksheetFunction.Max(Range("B:B").Value)
       'open_date = WorksheetFunction.Min(Cells(i, 2).Value)
       'close_date = WorksheetFunction.Max(Cells(i, 2).Value)
        close_price = WorksheetFunction.VLookup(close_date, Range("B:C").Value, 2, 0)
        open_price = WorksheetFunction.VLookup(open_date, Range("B:C").Value, 2, 0)
        
        'write ticker'
        Range("I" & New_Table).Value = Ticker
        Range("L" & New_Table).Value = Total_Stock_Volume
        Range("J" & New_Table).Value = close_price - open_price
        'write percent change'
        Range("K" & New_Table).Value = ((close_price - open_price) / open_price) * 100
    
        'add new row of data'
        New_Table = New_Table + 1

        'reset total'
        Total_Stock_Volume = 0

        Else
    
        'continue to add total stock volume'
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If
    Next i
'Next ws
End Sub

