Attribute VB_Name = "Module1"
Sub TestData()

Dim ws As Worksheet

For Each ws In Worksheets

' Declare datasheet variables

Dim ticker_symbol As String
Dim year_open As Double
Dim year_close As Double
Dim total_vol As Long
Dim open_price As Double
Dim close_price As Double
Dim year_change As Double
Dim percent_change As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Assign initial variables
year_open = ws.Range("C2").Value
total_vol = 0

'Assign Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Declare variabes for summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Create a loop that will output values to the summary table
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker_symbol = ws.Cells(i, 1).Value
year_close = ws.Cells(i, 5).Value

'Obtain difference between year close and year open
year_change = year_close - year_open

'Obtain percent change
If year_change = 0 Then
    percent_change = 0
Else
    percent_change = Round((year_change / year_open) * 100, 2)

End If

'Print values to summary table
ws.Cells(Summary_Table_Row, 9).Value = ticker_symbol
ws.Cells(Summary_Table_Row, 10).Value = year_change
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 12).Value = total_vol


'Add Rows
Summary_Table_Row = Summary_Table_Row + 1

'Reassign year open value for each ticker
year_open = ws.Cells(i + 1, 3).Value

Else
'Calculate total volume to include repeating ticker after initial ticker
total_vol = total_vol + ws.Cells(i, 6).Value

End If
    
Next i

Next ws
End Sub



