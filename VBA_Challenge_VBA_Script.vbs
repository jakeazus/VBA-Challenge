Sub Stock_Data()

Dim ws As Worksheet

'Loops through each worksheet in the excel document
For Each ws In Worksheets

'Defined variables
Dim Ticker_Symbol As String

Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Total_Volume As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Open_Price As Double
Open_Price = ws.Range("c2").Value

Dim Close_Price As Double

'finds the last row of data and return the index as an integer
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

'ws will represent each worksheet and will pull each .range or .cells into the worksheets

 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
 Ticker_Symbol = ws.Cells(i, 1).Value
 
 ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
 
 Close_Price = ws.Cells(i, 6).Value
 
 Yearly_Change = Close_Price - Open_Price
 
 ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
 
    If Open_Price = 0 Then
        'Prevents divide by zero error
        Percent_Change = 0
    Else
        Percent_Change = (Close_Price - Open_Price) / Open_Price
        
    End If
        
ws.Range("K" & Summary_Table_Row).Value = Percent_Change


    
'Adds total volume to the new column
Total_Volume = Total_Volume + ws.Cells(i, 7).Value
ws.Range("L" & Summary_Table_Row).Value = Total_Volume

Open_Price = ws.Cells(i + 1, 3).Value
 
Summary_Table_Row = Summary_Table_Row + 1

Total_Volume = 0
 
Else

Total_Volume = Total_Volume + ws.Cells(i, 7).Value

End If

Next i


For i = 2 To lastRow

If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
Else
    ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i


'Converts the percent change from decimal to percent
ws.Columns("K").NumberFormat = "0.00%"

'Prints the column title for each column
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"


Next ws

End Sub
