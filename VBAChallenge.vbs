Sub StockMarket():

'designate variables, multiple worksheets
Dim ws As Worksheet
Dim Ticker As String
Dim Ticker_Total As Double
'Dim LastRow As Long

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'designate headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
   
 'autofit column headers to fit values
    Columns("I:L").AutoFit
    
'setting total volume
    Total_volume = 0
    
'setting summary table to start at row 2 as row 1 is headers
    Summary_Table_Row = 2
    
'designate location of open price drawing data from for future calculations
    Open_Price = Cells(2, "C").Value

'designate last row to mark end of column data
    'LastRow = ws.Cells(ws.Rows.Count, "A").End(x1Up).Row - could not get this to run need TA help for scripting below
    LastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
'begin for loop to run data
    For i = 2 To LastRow
        
            Total_volume = Total_volume + Cells(i, "G").Value
 
    'designates when to change ticker value
            If ws.Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
        
    'designates closing price column
                    Close_price = Cells(i, "F").Value
            
    'formula used to determine yearly change
                    YearlyChange = Close_price - Open_Price
                
    'formula used to determine percent change
                    PercentChange = YearlyChange / Open_Price * 100
                
    'designate where to place calculations
                    ws.Cells(Summary_Table_Row, "I").Value = Cells(i, "A").Value
                
                    ws.Cells(Summary_Table_Row, "J").Value = YearlyChange
                
                    ws.Cells(Summary_Table_Row, "K").Value = PercentChange
                
                    ws.Cells(Summary_Table_Row, "L").Value = Total_volume
                
    'designate conditional formatting to yearly change column calculation
                    If YearlyChange > 0 Then
                        Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
                    
                    ElseIf YearlyChange < 0 Then
                        Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
                
                    Else
                        Cells(Summary_Table_Row, "J").Interior.ColorIndex = 2
                    
                    End If
                
    're-establish total values to re-start loop
                    Total_volume = 0
                
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    Open_Price = Cells(i + 1, "C").Value
        
            End If
            
        Next i
        
    Next ws
                
End Sub
