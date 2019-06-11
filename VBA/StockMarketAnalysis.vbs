Option Explicit
Sub StockMarket()

'Summing and recording total stock volumes for each ticker on each page

Dim sheet As Worksheet

For Each sheet In ActiveWorkbook.Worksheets

sheet.Activate

Dim ticker As String
Dim totalvol As Double
Dim openprice As Double
Dim yearlychange As Double
Dim percentchange As Double

'Declare the variable for the new table where the totals will be recorded:
Dim newTableRow As Double

Dim row As Double

newTableRow = 2
    
    'Initialize values for ticker and totalvol
    ticker = Range("A2").Value
    totalvol = Range("G2").Value
    openprice = Range("C2").Value

    'Set up the table headers for the totals for each ticker:
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim lastRow As Double
    
    lastRow = sheet.Cells(sheet.Rows.Count, "A").End(xlDown).row
    
    For row = 3 To lastRow
    
        'If the next ticker is the same as the previous ticker:
        If Cells(row, "A").Value = ticker Then
        
            totalvol = totalvol + Cells(row, "G").Value                 'adding to the total volume if ticker is the same
        
        
        'If the next ticker is different:
        Else: Cells(newTableRow, "I").Value = ticker                    'records the ticker symbol in our table for the totals
        
            Cells(newTableRow, "L").Value = totalvol                    'records the total stock volume for that ticker
        
        
        
            yearlychange = Cells(row - 1, "F").Value - openprice        'calculates yearly change
        
            Cells(newTableRow, "J").Value = yearlychange                'records the yearly change for that ticker
            
            If openprice = 0 Then
            percentchange = 0
        
            Else: percentchange = yearlychange / openprice              'calculates percent change
            
            End If
        
            Cells(newTableRow, "K").Value = percentchange       'records the percent change for that ticker
        
            Columns("K:K").Select
            Selection.Style = "Percent"                         'changes the percent change column to percentage formatting
        
            newTableRow = newTableRow + 1                       'pushes the new table row down one cell value for the next new ticker
        
            ticker = Cells(row, "A").Value                      'sets the new ticker as the value for the ticker variable
            
            totalvol = Cells(row, "G").Value                    'sets the new total volume as the value for the totalvol variable
        
            openprice = Cells(row, "C").Value                   'sets the new open price as the value for the openprice variable
        
        End If
        
    Next row
    
    'Conditional formatting for percent change
    For row = 2 To Rows.Count
        
            If Cells(row, "K").Value >= 0 Then
            Cells(row, "K").Interior.ColorIndex = 10
            
            ElseIf Cells(row, "K").Value < 0 Then
            Cells(row, "K").Interior.ColorIndex = 3
            
            End If
            
    Next row
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
'Find the greatest % increase, greatest % decrease, and greatest total stock volume for each sheet
    
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double
Dim GIticker As String
Dim GDecreaseticker As String
Dim GVticker As String


'Initialize values
greatestincrease = Range("K2").Value
greatestdecrease = Range("K2").Value
greatestvolume = Range("L2").Value
    
    For row = 3 To Rows.Count
        
        If Cells(row, "K").Value >= greatestincrease Then          'Find greatest percent increase
        
            greatestincrease = Cells(row, "K").Value                'Assign that value to the variable
            
            GIticker = Cells(row, "I").Value                        'Assign ticker to ticker variable
            
        End If
        
        
        If Cells(row, "K").Value <= greatestdecrease Then           'Find greatest percent decrease
        
            greatestdecrease = Cells(row, "K").Value                'Assign that value to the variable
            
            GDecreaseticker = Cells(row, "I").Value                 'Assign ticker to ticker variable
            
        End If
        
        
        If Cells(row, "L").Value >= greatestvolume Then             'Find greatest total volume
            
            greatestvolume = Cells(row, "L").Value                  'Assign that value to the variable
            
            GVticker = Cells(row, "I").Value                        'Assign ticker to ticker variable
            
        End If
        
    Next row
            
            
' Record values and tickers in table in each sheet
Range("P2").Value = GIticker
Range("Q2").Value = greatestincrease
Range("P3").Value = GDecreaseticker
Range("Q3").Value = greatestdecrease
Range("P4").Value = GVticker
Range("Q4").Value = greatestvolume

Range("Q2,Q3").Select
    Selection.Style = "Percent"
    
Next sheet


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Find greatest % increase, % decrease,and total stock volume values across all sheets

'Declare variables for the year/sheet
Dim yearGI As String
Dim yearGD As String
Dim yearGV As String

' Initialize values
 GIticker = Sheets("2016").Range("P2")
 greatestincrease = Sheets("2016").Range("Q2")
 GDecreaseticker = Sheets("2016").Range("P3")
 greatestdecrease = Sheets("2016").Range("Q3")
 GVticker = Sheets("2016").Range("P4")
 greatestvolume = Sheets("2016").Range("Q4")


For Each sheet In Worksheets

    sheet.Activate
     
    If Range("Q2").Value >= greatestincrease Then            'Find greatest percent increase value
    
        greatestincrease = Range("Q2").Value                'Assign that value to the variable
            
        GIticker = Range("P2").Value                        'Assign ticker to the ticker variable
        
        yearGI = sheet.Name                                 'Assign year to its variable
        
    End If
    
    
    If Range("Q3").Value <= greatestdecrease Then            'Find greatest percent decrease value
        
        greatestdecrease = Range("Q3").Value                'Assign that value to the variable
        
        GDecreaseticker = Range("P3").Value                 'Assign ticker to the ticker variable
        
        yearGD = sheet.Name                                 'Assign year to its variable
        
    End If
    
    
    If Range("Q4").Value >= greatestvolume Then              'Find greatest total volume value
    
        greatestvolume = Range("Q4").Value                  'Assign that value to the variable
        
        GVticker = Range("P4").Value                        'Assign ticker to the ticker variable
        
        yearGV = sheet.Name                                 'Assign year to its variable
        
    End If
    
Next sheet

' Create table in Sheet 2016
Worksheets("2016").Range("O8").Value = "Greatest values across all 3 years"
Worksheets("2016").Range("O10").Value = "Greatest % Increase"
Worksheets("2016").Range("O11").Value = "Greatest % Decrease"
Worksheets("2016").Range("O12").Value = "Greatest Total Volume"
Worksheets("2016").Range("P9").Value = "Ticker"
Worksheets("2016").Range("Q9").Value = "Value"
Worksheets("2016").Range("R9").Value = "Year"

' Record the values, tickers, and years in the table in Sheet 2016
Worksheets("2016").Range("Q10").Value = greatestincrease
Worksheets("2016").Range("P10").Value = GIticker
Worksheets("2016").Range("Q11").Value = greatestdecrease
Worksheets("2016").Range("P11").Value = GDecreaseticker
Worksheets("2016").Range("Q12").Value = greatestvolume
Worksheets("2016").Range("P12").Value = GVticker

Worksheets("2016").Range("R10").Value = yearGI
Worksheets("2016").Range("R11").Value = yearGD
Worksheets("2016").Range("R12").Value = yearGV

Range("Q10,Q11").Select
    Selection.Style = "Percent"

End Sub
