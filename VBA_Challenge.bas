Attribute VB_Name = "Module1"
Sub VBA_Challenge_Stock()

Dim Current As Worksheet

For Each Current In ThisWorkbook.Worksheets

'+++++++++++++++++
'Set and Bold Headers
'+++++++++++++++++
Range("I1") = "Ticker"
Range("i1").Font.Bold = True
Range("J1") = "Yearly Change"
Range("j1").Font.Bold = True
Range("K1") = "Percent Change"
Range("k1").Font.Bold = True
Range("L1") = "Total Stock Volume"
Range("L1").Font.Bold = True

'++++++++++++++
'Declare Variables
'++++++++++++++

Dim yearly_change As Double
Dim open_close_change As Double
Dim percent_change As Double
Dim total_stock As Double


Dim ticker As String
Dim ticker_total As Double
ticker_total = 0
Dim ticker_row As Integer
ticker_row = 2
open_close_change = 0
yearly_change = 2


'++++++++++++++++++++
'Loop through Ticker column
'++++++++++++++++++++



'Find last cell in colum A

Dim LRow As Long
    LRow = Cells(Rows.Count, 1).End(xlUp).Row

'Gather Tickers and Total Volumes
For i = 2 To LRow
    
        'check for matching tickers and add up total volume
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            'set the ticker name column
            ticker = Cells(i, 1).Value
            
            'add to the ticker total
            ticker_total = ticker_total + Cells(i, 7).Value
            
            'print the ticker  to the ticker column
            Range("I" & ticker_row).Value = ticker
            
            'print the ticker total to the total stock volume column
            Range("L" & ticker_row).Value = ticker_total
            
            'Add one to the ticker row
            ticker_row = ticker_row + 1
            
            'Reset the ticker total
            ticker_total = 0
            
        'check for yearly change per ticker
            
            'set the yearly_change column
            'yearly_change = Cells(i, 10).Value
            
            'gather the yearly change averages
            'Application.WorksheetFunction.Average(Range
            
            'print  the yearly change to the column
            'Range("J" & yearly_change).Value = open_close_change
            
            'add 1 to the yearly_change
            'yearly_change = yearly_change + 1
            
            'reset the open_close_change
            'open_close_change = 0
            
            Else
            
            'add to the ticker total
            ticker_total = ticker_total + Cells(i, 7).Value


    End If
     
Next i

Next Current


End Sub

