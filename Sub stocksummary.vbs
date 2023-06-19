Sub stocksummary()
'loop through all sheets
For Each ws In Worksheets

'set number of rows as long, stockname (ticker) as string, yearly price change as double
'checked for a solution to the error "Overflow (Error 6)"  changed totalstockvolume to double and it run.  Not sure why.  need to check.
Dim lastrow As Long
Dim ticker As String
Dim yearlypricechange As Double
'set percent change as double, total stock volume as long, date as long, open value as double
Dim percentchange As Double
Dim totalstockvolume As Double
Dim datenumber As Long
Dim openvalue As Double
'set close value as double, volume as long
Dim closevalue As Double
Dim volume As Long

'set values for counters:
totalstockvolume = 0

'set values for summary table, ticker, openvalue
Dim summaryrow As Long
summaryrow = 2
ticker = ws.Cells(summaryrow, 1).Value
openvalue = ws.Cells(summaryrow, 3).Value

'set titles of new columns:
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'count the number of rows
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through all the stocks sales
For i = 2 To lastrow
'check when a stock changes name
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'write the ticker in the summary table
ws.Cells(summaryrow, 9).Value = ticker

'calculate the change in price
closevalue = ws.Cells(i, 6).Value
yearlypricechange = closevalue - openvalue

'write the yearly price change in the summary table
ws.Cells(summaryrow, 10).Value = yearlypricechange

'calculate the percent change
percentchange = yearlypricechange / openvalue

'write the yearly percent change in the summary table
ws.Cells(summaryrow, 11).Value = percentchange

'write the yearly volume in the summary table
ws.Cells(summaryrow, 12).Value = totalstockvolume

'update values for the next loop
summaryrow = summaryrow + 1
ticker = ws.Cells(i + 1, 1).Value
openvalue = ws.Cells(i + 1, 3).Value
totalstockvolume = 0
End If

'Add the volume of transactions for the stock
totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value   'not adding the last volume of each stock

Next i

'set values for summary table2
Dim summary2row As Long
summary2row = 2
Dim colrangepercent As Range
Dim colrangevolume As Range
Dim greatestpercentincrease As Double
Dim greatestpercentdecrease As Double
Dim greatestvolume As Double
Dim lastrowsummarytable1 As Long
Dim tickersummary2 As String

'set titles of new columns in summary table 2:
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"

'count the number of rows in summary table 1
lastrowsummarytable1 = ws.Cells(Rows.Count, "I").End(xlUp).Row

' Set the range for the column in the summary table 1 with percent changes
Set colrangepercent = ws.Range("K1:K" & lastrowsummarytable1)

' Use the WorksheetFunction.Max function to get the maximum value (from documentation)
greatestpercentincrease = Application.WorksheetFunction.Max(colrangepercent)

'write the max percent increase in the summary table 2
ws.Range("q2").Value = greatestpercentincrease

' Use the WorksheetFunction.Min function to get the minimum value (from documentation)
greatestpercentdecrease = Application.WorksheetFunction.Min(colrangepercent)

'write the max percent decrease in the summary table 2
ws.Range("q3").Value = greatestpercentdecrease

' Set the range for the column in the summary table 1 with volume
Set colrangevolume = ws.Range("L1:L" & lastrowsummarytable1)

' Use the WorksheetFunction.Max function to get the maximum volume(from documentation)
greatestvolume = Application.WorksheetFunction.Max(colrangevolume)

'write the highest volume in the summary table 2
ws.Range("q4").Value = greatestvolume

'loop to compare values in summary table 2 with summary table 1
For summaryrow = 2 To lastrowsummarytable1

'conditional to set ticker values for greatest percent increase in summary tabl 2
If ws.Range("q2") = ws.Cells(summaryrow, "I") Then

'set value for ticker 2 in summary table 2
tickersummary2 = ws.Cells(summaryrow, "I").Value

'write ticker in summary table 2 for
ws.Cells(2, "P").Value = tickersummary2

End If

Next summaryrow

Next ws

End Sub

