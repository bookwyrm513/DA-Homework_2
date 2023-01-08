Sub stocktick()

'initialize variables for program
Dim tickstock As Integer
Dim opentotal As Double
Dim closetotal As Double
Dim totalvol As Double
Dim topinc As Double
Dim toptick As String
Dim topdec As Double
Dim lowtick As String
Dim topvol As Double
Dim voltick As String
Dim lastrow As Long
Dim lastrow2 As Long
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
'telling the program to start here
tickstock = 2
opentotal = ws.Cells(2, 3).Value
closetotal = ws.Cells(2, 6).Value
totalvol = ws.Cells(2, 7).Value
topinc = ws.Cells(2, 11).Value
topdec = ws.Cells(2, 11).Value
topvol = ws.Cells(2, 12).Value
'determine last row, code taken from (https://stackoverflow.com/questions/45510730/vba-how-to-convert-a-column-to-percentages)
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
'change format of column k (not including header) to percentage format, taken from same code source as above
ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
'formatting for if there is no header present
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"


'iterate through entire sheet and tally up totals for individual stocks
For i = 2 To lastrow
    'check if next cell ticker matches this cell ticker
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        'tally up the next cell open value
        opentotal = opentotal + ws.Cells(i + 1, 3).Value
        'tally up the next cell close value
        closetotal = closetotal + ws.Cells(i + 1, 6).Value
        'tally the volumes
        totalvol = totalvol + ws.Cells(i + 1, 7).Value
    'else there is a change in stock ticker
    Else
        'output stock name
        ws.Cells(tickstock, 9).Value = ws.Cells(i, 1).Value
        'output yearly change
        ws.Cells(tickstock, 10).Value = closetotal - opentotal
        'change color based on value of yearly change
        'please note that I can in fact do this with conditional formatting within the sheet itself
        'this is just faster and easier
        If closetotal - opentotal <= 0 Then
            ws.Cells(tickstock, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(tickstock, 10).Interior.ColorIndex = 4
        End If
        'percent change
        ws.Cells(tickstock, 11).Value = (closetotal - opentotal) / closetotal * 100
        'total stock volume
        ws.Cells(tickstock, 12).Value = totalvol
        'reset counters to next stock
        tickstock = tickstock + 1
        opentotal = ws.Cells(i + 1, 3).Value
        closetotal = ws.Cells(i + 1, 6).Value
        totalvol = ws.Cells(i + 1, 7).Value
    End If
Next i



'determine last row of created chart, raw code taken from (https://stackoverflow.com/questions/45510730/vba-how-to-convert-a-column-to-percentages)
lastrow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
'change format of P2:P3 (not including header) to percentage format, taken from same code source as above
ws.Range("P2:P3").NumberFormat = "0.00%"

'search the created chart to find greatest % increase, decrease and total volume
For j = 2 To lastrow2
    'compare current top values and current values and determine larger/smaller value
    If ws.Cells(j, 11).Value > topinc Then
        'set top $ increase to current value and output the stock and value
        topinc = ws.Cells(j, 11).Value
        toptick = ws.Cells(j, 9).Value
    ElseIf Cells(j, 11).Value < topdec Then
        'same as above but with lowest value
        topdec = ws.Cells(j, 11).Value
        lowtick = ws.Cells(j, 9).Value
    End If
    'switch to checking if current cell
    If ws.Cells(j, 12).Value > topvol Then
        topvol = ws.Cells(j, 12).Value
        voltick = ws.Cells(j, 9).Value
    End If
Next j
    'display output
    ws.Cells(2, 15).Value = toptick
    ws.Cells(2, 16).Value = topinc
    ws.Cells(3, 15).Value = lowtick
    ws.Cells(3, 16).Value = topdec
    ws.Cells(4, 15).Value = voltick
    ws.Cells(4, 16).Value = topvol

'autofit the sheet to look nicer
ws.Range("A:P").EntireColumn.AutoFit

Next

End Sub
