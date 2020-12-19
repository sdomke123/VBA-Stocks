Attribute VB_Name = "Module1"
Sub StockData()

'Apply code for all worksheets
For Each ws In Worksheets

'Define all variables used
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim vol As Long
Dim voltotal As LongLong
Dim yearchange As Double
Dim percent As Double
Dim i As Long
Dim Summary_Row As Long
Dim FirstTickerRow As Long
Dim LastTickerRow As Long

'Set some initial values
voltotal = 0
Summary_Row = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
openprice = ws.Cells(2, 3)

'Display summary headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Loop through all rows, setting conditions for when a ticker value is different than the next
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            closeprice = ws.Cells(i, 6).Value
            voltotal = voltotal + ws.Cells(i, 7).Value
            yearchange = closeprice - openprice
            
        'Disregard when openprice is 0 to avoid dividing by 0
        If openprice = 0 Then
            percent = 0
        Else
            percent = (yearchange / openprice)
        End If
                  
            'Display the resulting values in corresponding ranges
            ws.Range("I" & Summary_Row).Value = ticker
            ws.Range("J" & Summary_Row).Value = yearchange
            ws.Range("K" & Summary_Row).Value = percent
            ws.Range("L" & Summary_Row).Value = voltotal
            
            'Set conditions for following ticker loops
            Summary_Row = Summary_Row + 1
            voltotal = 0
            openprice = ws.Cells(i + 1, 3)
            
     Else
        voltotal = voltotal + ws.Cells(i, 7).Value
            
    End If
    
    'Fill positive yearly changes in green, negative in red
    If ws.Range("J" & Summary_Row).Value < 0 Then
        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
        
    ElseIf ws.Range("J" & Summary_Row).Value > 0 Then
        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
        
    End If

Next i
   
Next ws



End Sub

