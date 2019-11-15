Sub yearly_summary()
    
    ' # READ ME *** | Script logic assumes the following
    ' # - The Ticker names (Column A:A) are in alphabetical order
    ' # - The Dates (Column B:B) are in chronilogical order, from oldest to newest.

    'ALL DEFINED VARIABLES --------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Volume_Total As Double
    Dim Yearly_Difference As Double
    Dim Yearly_Percent_Change As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    'All numerical variables start value of 0
    Volume_Total = 0
    Yearly_Difference = 0
    Yearly_Percent_Change = 0
    Open_Price = 0
    Close_Price = 0
    
    'EACH WORKSHEET --------------------------------------------------------------------------
    Set Workbook = ActiveWorkbook
    For Each ws In Workbook.Worksheets
    ws.Activate
        'Find last row of current sheet
        Dim LastRow As Long
        LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
        'Summary Table Row Variable
        Dim summary_row As Integer
        summary_row = 2
      
    'RUN ALL DATA ROWS - LOOP --------------------------------------------------------------------------
        For i = 2 To LastRow
        
            'FIND FIRST ROW OF GROUP
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Define Ticker name
                    Ticker = ws.Cells(i, 1).Value
                    'Define Open price
                    Open_Price = ws.Cells(i, 3).Value
            'FIND ALL ROWS INSIDE OF GROUP
                ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                    'Add to Volume_Total
                    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            'FIND LAST ROW OF GROUP
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Define Close Price
                    Close_Price = ws.Cells(i, 6).Value
                    'Find Yearly Difference
                    Yearly_Difference = Close_Price - Open_Price
                    'Find Yearly Percent Change
                            'If Percent Increase
                            If Close_Price > Open_Price Then
                            Yearly_Percent_Change = (Close_Price - Open_Price) / Close_Price
                            'If Percent Decrease
                            ElseIf Open_Price > Close_Price Then
                            Yearly_Percent_Change = ((Open_Price - Close_Price) / Open_Price) * -1
                            End If
                    'Add volume to total
                    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
            'PRINT RESULTS
                ws.Range("I" & summary_row).Value = Ticker
                ws.Range("J" & summary_row).Value = Yearly_Difference
                        'Conditional Formatting if Increase
                        If ws.Range("J" & summary_row).Value >= 0 Then
                        ws.Range("J" & summary_row).Interior.ColorIndex = 43
                        Else 'If decrease
                        ws.Range("J" & summary_row).Interior.ColorIndex = 53
                        End If
                ws.Range("K" & summary_row).Value = Yearly_Percent_Change
                ws.Range("L" & summary_row).Value = Volume_Total
            
            'CLEAR FOR NEXT LOOP
                'Add next row to summary table
                summary_row = summary_row + 1
                'Reset Variables
                Open_Price = 0
                Close_Price = 0
                Yearly_Difference = 0
                Yearly_Percent_Change = 0
                Volume_Total = 0
                
            End If
                
        Next i
        
    'FINISH SUMMARY TABLE --------------------------------------------------------------------------
        'Create Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Difference"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        'Apply Formatting
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        
    'CREATE SUMMARY - OF - SUMMARY --------------------------------------------------------------------------
        'Define summary of summary variables
            Dim MaxName As String
            Dim MinName As String
            Dim VolName As String
            Dim MaxRow As Double
            Dim MinRow As Double
            Dim VolRow As Double
            MaxRow = 0
            MinRow = 0
            VolRow = 0
            'Find last row of summary table
            Dim LastSummaryRow As Long
            LastSummaryRow = ActiveSheet.Range("I" & Rows.Count).End(xlUp).Row
            
        'Begin looping through summary table
            For s = 2 To LastSummaryRow
                'Find the Max % increase
                    If ws.Cells(s, 11).Value >= MaxRow Then
                        MaxRow = ws.Cells(s, 11).Value
                        MaxName = ws.Cells(s, 9).Value
                'Find the Min % increase
                    ElseIf ws.Cells(s, 11).Value <= MinRow Then
                        MinRow = ws.Cells(s, 11).Value
                        MinName = ws.Cells(s, 9).Value
                    End If
                    
                'Find Max Total Volume
                    If ws.Cells(s, 12).Value >= VolRow Then
                        VolRow = ws.Cells(s, 12).Value
                        VolName = ws.Cells(s, 9).Value
                    End If
                    
            Next s
        'FINISH SUMMARY - OF - SUMMARY TABLE --------------------------------------------------------------------------
            'Print Results found:
                ws.Cells(2, 16).Value = MaxRow
                ws.Cells(2, 15).Value = MaxName
                ws.Cells(3, 16).Value = MinRow
                ws.Cells(3, 15).Value = MinName
                ws.Cells(4, 16).Value = VolRow
                ws.Cells(4, 15).Value = VolName
            'Create Headers
                ws.Cells(1, 14).Value = "Outliers"
                ws.Cells(2, 14).Value = "Greatest % Increase"
                ws.Cells(3, 14).Value = "Greatest % Decrease"
                ws.Cells(4, 14).Value = "Greatest Total Volume"
                ws.Cells(1, 15).Value = "Ticker"
                ws.Cells(1, 16).Value = "Value"
            'Apply Formatting
                ws.Range("N1:P1").Font.Bold = True
                ws.Range("P2:P3").NumberFormat = "0.00%"
                ws.Columns("N:P").AutoFit
            
    Next ws
    
    'SCRIPT DONE, RETURN TO FIRST WORKSHEET --------------------------------------------------------------------------
    Application.Goto Reference:=Worksheets(1).Range("A1"), _
        Scroll:=True
    MsgBox ("Hooray! Your summary tables have now been generated. Have a great day!")

End Sub
