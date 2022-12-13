Attribute VB_Name = "Module1"
Sub main():
'For each sheet, we'll want to calculate the yearly change/percent change/total stock volume for each ticker
'At the broadest level, we'll need to iterate through the sheets, then we'll want to do all of our calculations
'first, we need to figure out how many sheets there are
Dim wsCnt As Integer
Dim wsCurrent As Integer

wsCnt = ActiveWorkbook.Worksheets.Count
    
'for each sheet in the workbook, do all the things
For wsCurrent = 1 To wsCnt
    Worksheets(wsCurrent).Select
    evaluateChanges
    conditionalFormat
    evaluateGreatests
    Worksheets(wsCurrent).Cells.EntireColumn.AutoFit
Next wsCurrent
End Sub
Sub evaluateChanges():
    Dim ticker As String
    Dim tickerFirstRow As Double
    Dim priceOpen As Double
    Dim priceClose As Double
    Dim totalVol As Double
    Dim tableCounter As Double
    Dim lastRow As Double
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'last populated row in the current sheet
    'we need to initialize the first row for the first ticker to 2. it will change later on.
    tickerFirstRow = 2
    tableCounter = 2 'this is the row counter for the output table
    
    'set header row
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'grab the values needed for the current ticker
            ticker = Cells(i, 1).Value
            priceOpen = Cells(tickerFirstRow, 3).Value
            priceClose = Cells(i, 6).Value
            totalVol = totalVol + Cells(i, 7).Value
            
            'set the information into the table
            Cells(tableCounter, 9).Value = ticker
            Cells(tableCounter, 10).Value = priceClose - priceOpen
            Cells(tableCounter, 11).Value = Format((priceClose - priceOpen) / priceOpen, "0.00%")
            Cells(tableCounter, 12).Value = totalVol
            
            'iterate the tableCounter, change the next ticker first row, and reset the volume total
            tableCounter = tableCounter + 1
            tickerFirstRow = i + 1
            totalVol = 0
        'else we are evaluating the same ticker and just need to add the volume to the total
        Else
            totalVol = totalVol + Cells(i, 7).Value
        'onto the next!
        End If
    Next i
End Sub
Sub conditionalFormat():
    Dim lastRow As Integer
    
    lastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
    'evaluate if the current row ticker's Yearly Change is negative or not
    'if it's negative, highlight the cell in red (color index 3), if not, highlight the cell in green (color index 4)
    For i = 2 To lastRow
        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
    
End Sub
Sub evaluateGreatests():
    Dim tickerMax As String
    Dim tickerMin As String
    Dim tickerVol As String
    Dim rowCnt As Integer
    Dim lastRow As Integer
    Dim max As Double
    Dim min As Double
    Dim vol As Double
    
    'set labels/headers
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    lastRow = Cells(Rows.Count, 11).End(xlUp).Row
    max = -99999
    min = 99999
    vol = -1
    
    For rowCnt = 2 To lastRow
        If Cells(rowCnt, 11).Value > max Then
            tickerMax = Cells(rowCnt, 9).Value
            max = Cells(rowCnt, 11).Value
        End If
        If Cells(rowCnt, 11).Value < min Then
            tickerMin = Cells(rowCnt, 9).Value
            min = Cells(rowCnt, 11).Value
        End If
        If Cells(rowCnt, 12).Value > vol Then
            tickerVol = Cells(rowCnt, 9).Value
            vol = Cells(rowCnt, 12).Value
        End If
    Next rowCnt
    
    Cells(2, 16).Value = tickerMax
    Cells(2, 17).Value = max
    Cells(3, 16).Value = tickerMin
    Cells(3, 17).Value = min
    Cells(4, 16).Value = tickerVol
    Cells(4, 17).Value = vol
    
End Sub

