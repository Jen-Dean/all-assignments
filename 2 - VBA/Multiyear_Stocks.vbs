Attribute VB_Name = "Module1"
Sub AllSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stocks
    Next
    Application.ScreenUpdating = True
End Sub

Sub Stocks():

'Set Up Labels on All Sheets

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("N2").Value = "Greatest Increase"
Range("N3").Value = "Greatest Decrease"
Range("N4").Value = "Greatest Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'Define Variables'

Dim Ticker As String

Dim Total As Double
        Total = 0
        
Dim OpenPrice As Double
        OpenPrice = 0
        
Dim ClosePrice As Double
        ClosePrice = 0
        
Dim YearlyChange As Double
        
Dim PercentChange As Double
        
Dim TableRow As Double
        TableRow = 2
        
Dim GI As Double
        GI = 0
Dim GITicker As String
     
Dim GD As Double
        GD = 0
Dim GDTicker As String
        
Dim GV As Double
        GV = 0
Dim GVTicker As String

'Find Last Row in the Sheet'
Dim MaxRow As Double
        MaxRow = ActiveSheet.UsedRange.Rows.Count

'Set the Initial Open Price'
OpenPrice = Cells(2, 3).Value
'         Cells(TableRow, 11).Value = OpenPrice < This was used for testing.

'Set a Loop to go through all rows - to MaxRow'
For i = 2 To MaxRow

        If OpenPrice = 0 Then
            GoTo NextTicker
        End If
        
'Determine when the Loop should stop to reset the data, or keep going'
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                Total = Total + Cells(i, 7).Value
'What should be recorded when finding a new ticker symbol'
                Else
                        Total = Total + Cells(i, 7).Value
                        Cells(TableRow, 12).Value = Total
                        
                        Ticker = Cells(i, 1).Value
                        Cells(TableRow, 9).Value = Ticker
                        
                        ClosePrice = Cells(i, 6).Value
                        
                        YearlyChange = (OpenPrice - ClosePrice)
                        Cells(TableRow, 10).Value = YearlyChange
                        
                        PercentChange = ((OpenPrice - ClosePrice) / OpenPrice)
                        Cells(TableRow, 11).Value = FormatPercent(PercentChange, [-1])
                        
                        OpenPrice = Cells(i + 1, 3).Value
                        Cells(TableRow + 1, 11).Value = OpenPrice
 
'Conditional Formating - Yearly Change'
                        If Cells(TableRow, 10).Value < 0 Then
                                Cells(TableRow, 10).Interior.ColorIndex = 3
                        Else
                                Cells(TableRow, 10).Interior.ColorIndex = 4
                        End If
'Conditional Formating - Percent Change'
                        'If Cells(TableRow, 11).Value < 0 Then
                                'Cells(TableRow, 11).Interior.ColorIndex = 3
                        'Else
                                'Cells(TableRow, 11).Interior.ColorIndex = 4
                        'End If
                        
'Reset the variables before starting the next loop'
                        Total = 0
                        YearlyChange = 0
                        TableRow = TableRow + 1
        End If
NextTicker:
Next i

'Fine Last Row in New Table Columns
Dim iRow As Double
        iRow = Cells(Rows.Count, 9).End(xlUp).Row

'Find the Greatest % Increase`
For i = 2 To iRow
       If Cells(i, 10).Value > GI Then
            GI = Cells(i, 10).Value
            GITicker = Cells(i, 9).Value
       End If
        
        If Cells(i, 10).Value < GD Then
            GD = Cells(i, 10).Value
            GDTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > GV Then
           GV = Cells(i, 12).Value
           GVTicker = Cells(i, 9).Value
        End If
Next i

'Print the results
Range("P2").Value = GI
Range("O2").Value = GITicker
Range("P3").Value = GD
Range("O3").Value = GDTicker
Range("P4").Value = GV
Range("O4").Value = GVTicker

End Sub
