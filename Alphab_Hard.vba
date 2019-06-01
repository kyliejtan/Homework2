Sub Greatest_Calc()
        'Setting the variable to hold each worksheet in the workbook
        Dim ws As Worksheet
        'Looping through all worksheets
        For Each ws In Worksheets
            'Setting the sheet that the outer loop is currently working on the active sheet
            ws.Activate
            'Setting variables to hold the Lastrow, Greatest... categories, and their respective cells
            Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            Dim Greatest_Increase As Double
            Greatest_Increase = 0
            Range("Q2").Value = Greatest_Increase
            Dim Greatest_Decrease As Double
            Greatest_Decrease = 0
            Range("Q3").Value = Greatest_Decrease
            Dim Greatest_Total_Volume As Double
            Greatest_Total_Volume = 0
            Range("Q4").Value = Greatest_Total_Volume
            'Printing Ticker, Value, and Greatest... category headers to their respective columns
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            'Making the Ticker, Value, and Greatest category headers fonts bold
            Range("P1").Font.Bold = True
            Range("Q1").Font.Bold = True
            Range("O2").Font.Bold = True
            Range("O3").Font.Bold = True
            Range("O4").Font.Bold = True
            'Auto fitting the column width for the Ticker, Value, and Greatest category headers
            Range("P1").EntireColumn.AutoFit
            Range("Q1").EntireColumn.AutoFit
            Range("O2").EntireColumn.AutoFit
            Range("O3").EntireColumn.AutoFit
            Range("O4").EntireColumn.AutoFit
                'Looping through each row of the WS that the outer loop is currently working on
                For i = 2 To Lastrow
                    'Check if the current row that the inner loop is working Is greater than the last value stored for each category
                    If Cells(i, 11).Value > Greatest_Increase Then
                        'Set the Ticker Name
                        Range("P2").Value = Cells(i, 9).Value
                        'Amending the Greatest Increase if possible
                        Greatest_Increase = Cells(i, 11).Value
                    ElseIf Cells(i, 11).Value < Greatest_Decrease Then
                        'Set the Ticker Name
                        Range("P3").Value = Cells(i, 9).Value
                        'Amending the Greatest Decrease if possible
                        Greatest_Decrease = Cells(i, 11).Value
                    ElseIf Cells(i, 12).Value > Greatest_Total_Volume Then
                        'Set the Ticker Name
                        Range("P4").Value = Cells(i, 9).Value
                        'Amending the GreatestTotal Volume if possible
                        Greatest_Total_Volume = Cells(i, 12).Value
                    Else
                        
                    End If
                Next i
                'Printing the final values for each category to their respective cells
                Range("Q2").Value = Greatest_Increase
                Range("Q3").Value = Greatest_Decrease
                Range("Q4").Value = Greatest_Total_Volume
            Next ws
    End Sub