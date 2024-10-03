VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()
   
    Dim i As Long
    Dim j As Double
    Dim ws As Worksheet
        
'Loop through all sheets
    For Each ws In Worksheets
    
        'variables
        Dim LastRow As Long
        
        Dim ticker_name As String
        Dim ticker_total As Double
        ticker_total = 0
    
        'add a column header for ticker and volume, output ticker symbol(once) and total volume
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'insert column and row names
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Keep track of ticker name, price change, and volume totals
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        'Keep track of first opening price and last closing price
        Dim opening_price As Double
        Dim closing_price As Double
        Dim quarterly_change As Double
               
    
        'Keep track of min, max and total values
        'Dim max As Double
        'Dim min As Double
        'Dim max_volume As Long
        
    'create a script that loops through all stocks for each quarter
                
        For i = 2 To LastRow
                             
        'check if new ticker name
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'store opening price of the new ticker at the beginning of the quarter
                opening_price = ws.Cells(i, 3).Value
                
            End If
                               
        'check for last row of current ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            
                'set ticker new name
                ticker_name = ws.Cells(i, 1).Value
                                               
                'add to ticker volume total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                
                'store closing price of previous ticker at the end of the quarter
                closing_price = ws.Cells(i, 6).Value
                                            
                'calculate and print quarterly change
                quarterly_change = closing_price - opening_price
                ws.Range("J" & Summary_Table_Row).Value = quarterly_change
                
                'conditional formatting Column J Quarterly change positive in green and negative in red
                                
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    ElseIf Range("J" & Summary_Table_Row).Value < 0 Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    End If
                
                'reset opening price and print quarterly change as percent
                If opening_price = 0 Then
            
                    'store opening price once
                    opening_price = ws.Cells(i, 3).Value
                    
                ElseIf opening_price <> 0 Then
                
                    'Calculate and print quarterly change as percent
                    ws.Range("K" & Summary_Table_Row).Value = (quarterly_change / opening_price)
                    
                    'format cells to percent
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                                        
                Else
                
                    ws.Range("K" & Summary_Table_Row).Value = 0
                    
                End If
                    
                'print the ticker name to the summary table
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
        
                'print the ticker (stock volume) total to the summary table
                ws.Range("L" & Summary_Table_Row).Value = ticker_total
        
                'add one to the summary table
                Summary_Table_Row = Summary_Table_Row + 1
            
                'reset the ticker total
                ticker_total = 0
            
            Else
                                
                'keep adding to the ticker total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                                
            End If
            
        Next i
        
        'identify greatest increase in Column K Percent Change
        ' take the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
                    
        ' returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = Cells(increase_number + 1, 9)
        ws.Range("P3") = Cells(decrease_number + 1, 9)
        ws.Range("P4") = Cells(volume_number + 1, 9)
            
    ws.Columns("I:Q").AutoFit
        
    Next ws
           
End Sub

