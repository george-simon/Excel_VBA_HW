Attribute VB_Name = "Module1"
Sub alphabetical()
    
        'Define ticker name
        Dim ticker As String
        
        'Set up variable to hold volume
        Dim volume As Variant
        volume = 0
    
        'Set up variables
        Dim yearly_open As Double
        Dim yearly_close As Double
        Dim percent_change As Double
        Dim yearly_change As Double
        
        'Summary table location
        Dim summarytable_row As Integer
        summarytable_row = 2
        
        'Define last row & table row
        Dim last_row As Variant
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        'prevents my overflow error: https://stackoverflow.com/questions/2202869/what-does-the-on-error-resume-next-statement-do
        'This gets around my overflow error for dividing by zero
        On Error Resume Next
        
        'next_ws variable
        'Cycling through worksheets: https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
        Dim ws As Worksheet
        'Dim next_ws As Integer
        'next_ws = 7
        
    
    'loop through all worksheets - insight: it looks to be anywhere there are 'cells' in y code I would add ws
    'NOTE - I keep getting an overflow error when I add <ws.> before the cell syntax in the nested loop...
    For Each ws In Worksheets
    'For ws = 1 To next_ws
    
    'insert headers for summary table
    ws.Cells(1, 10).Value = "ticker"
    ws.Cells(1, 11).Value = "yearly_change"
    ws.Cells(1, 12).Value = "percent_change"
    ws.Cells(1, 13).Value = "volume"
    
        'Set loop for all stocks
        For r = 2 To last_row
        'last_row
            'Check to see if stock ticker in the first row
            If ws.Cells(r, 1).Value <> ws.Cells(r - 1, 1).Value Then
            
                'Set the ticker, volume, and year open
                ticker = ws.Cells(r, 1).Value
                volume = ws.Cells(r, 7).Value
                year_open = ws.Cells(r, 3).Value
                year_close = ws.Cells(r, 6).Value
                
            End If
            
            ' Check to see if the stock ticker is the same in the last row
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                'set close value
                year_close = ws.Cells(r, 6).Value
                
                'add to the volume
                volume = volume + ws.Cells(r, 7).Value
                
                'percent and yearly change calculation
                percent_change = ((year_close - year_open) / year_open)
                yearly_change = year_close - year_open
          
            
                'Print to summary table
                ws.Cells(summarytable_row, 13).Value = volume
                ws.Cells(summarytable_row, 12).Value = percent_change
                ws.Cells(summarytable_row, 10).Value = ticker
                ws.Cells(summarytable_row, 11).Value = yearly_change
                
                'Add one to the summary table row
                summarytable_row = summarytable_row + 1
                
                'reset volume to zero
                volume = 0
            
            'if the ticker is the same
            Else
                'add to volume
                volume = volume + ws.Cells(r, 7).Value
        
            
            End If
           
        Next r
        
        'Format column L for percent
        Columns("L").NumberFormat = "0.00%"
        
        'Define format variables
        Dim compl_range As Variant
        compl_range = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Loop through column K for + / - yearly change
        For f = 2 To compl_range
            If ws.Cells(f, 11).Value >= 0 Then
            'Color green for positive
            ws.Cells(f, 11).Interior.ColorIndex = 4
        
            ElseIf ws.Cells(f, 11).Value < 0 Then
            'Color red for negative
            ws.Cells(f, 11).Interior.ColorIndex = 3
        
            End If

        Next f
        
    Next ws
    
End Sub
    
