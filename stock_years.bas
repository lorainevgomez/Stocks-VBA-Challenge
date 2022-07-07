Attribute VB_Name = "stock_years"
Sub stock_years()


'Defining Variables for Worksheet, Range, and LastRow
Dim ws As Worksheet
Dim Rng As Range
Dim LastRow As Long
Dim i As Long



'Set initial variable for holding total per Ticker
Dim stockvolume As Double
stockvolume = 0
Dim Ticker As String


'Defining variables for Yearly Change
Dim open_year As Double
Dim close_year As Double
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double



'Runs through all work sheets
For Each ws In Worksheets
    'Define last row of A
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'Adjust columns
    ws.Columns("A:M").AutoFit
    
    'Set Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"



        'Keep track of location of each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        'Keep track of location of each dates open price
        Dim price_counter As Long
        price_counter = 2
    
                'Loop through all the Tickers
                For i = 2 To LastRow
                
                
    
                        
    
                        'Loop to check if we are in the same Ticker, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                        'Find Ticker
                        Ticker = ws.Cells(i, 1).Value
                        'Add to the total volume
                        stockvolume = stockvolume + ws.Cells(i, 7).Value
    
                        'print tickers into the summary table
                        ws.Range("I" & Summary_Table_Row).Value = Ticker
                        
                        
                        'print the total amount into summary table
                        ws.Range("L" & Summary_Table_Row).Value = stockvolume
                        
                        'Find yearly change
                         'value of open year
                        open_year = ws.Range("C" & price_counter).Value
                        
                        close_year = ws.Range("F" & i).Value
                        yearly_change = close_year - open_year
                        
                            If open_year = 0 Then
                                percent_change = 0
                                
                            Else
                                percent_change = yearly_change / open_year
                                
                            End If
                            
                        
                        'print the yearly change into summary table
                        ws.Range("J" & Summary_Table_Row).Value = yearly_change
                        
                        'print percent change into summary table
                        
                        ws.Range("K" & Summary_Table_Row).Value = percent_change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            
                        'Add one to the summary table row
                        Summary_Table_Row = Summary_Table_Row + 1
                        price_counter = i + 1
    
                        'reset
                        stockvolume = 0
                        
            
                        'If the cell immediately following a row is the same ticker
                 Else
        
                        'Add to the volume total
                        stockvolume = stockvolume + ws.Cells(i, 7).Value
                        
                        
        
                 End If
                 
                 
                        'Conditional Formatting for yealy change (Negative is Red and Green is Positive)
                        
                        If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                            
                                    ElseIf ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                        
                                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                            
                                    ElseIf ws.Cells(Summary_Table_Row, 10).Value = 0 Then
                        
                                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 0
                            
                        End If
                        
                        
    Next i

    

Next ws


End Sub
