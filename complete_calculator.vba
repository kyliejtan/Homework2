Sub Complete_Calculator()
    'This calculator is broken up into three sub-routines. When Completel_Calculator() is \
    'finished, it calls Yearly_And_Percentage_Change_Calc() which in turn calls \
    'Greatest_Calc() when it is finished. These three sub-routines make up the all of the \
    'necessary code to complete the Homework 2 assignment at the hard level. 
    'Setting the variable to hold each worksheet in the workbook
    Dim ws As Worksheet
    'Looping through all worksheets
    For Each ws In Worksheets
        'Setting the sheet that the outer loop is currently working on the active sheet
        ws.Activate
        'Setting variables to hold the Lastrow, Ticker Name, Total Volume, and space the ticker names in the summary table.
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Dim Ticker_Name As String
        Dim Total_Volume As Double
        Total_Volume = 0
        Dim Summary_Table_Row_Spacer As Integer
        Summary_Table_Row_Spacer = 2
        'Printing Ticker and Total Stock Volume headers to their respective columns
        Range("I1").Value = "Ticker"
        Range("L1").Value = "Total Stock Volume"
        'Making the Ticker and Total Stock Volume hearder fonts bold
        Range("I1").Font.Bold = True
        Range("L1").Font.Bold = True
        'Auto fitting the column with of the Ticker and Total Stock Volume columns
        Range("I1").EntireColumn.AutoFit
        Range("L1").EntireColumn.AutoFit
            'Looping through each row of the WS that the outer loop is currently working on
            For i = 2 To Lastrow
                'Check if the current row that the inner loop is working on does not match the next row
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'Set the Ticker Name
                    Ticker_Name = Cells(i, 1).Value
                    'Add to the Total Volume
                    Total_Volume = Total_Volume + Cells(i, 7).Value
                    'Print the Ticker Name to the summary table
                    Range("I" & Summary_Table_Row_Spacer).Value = Ticker_Name
                    'Print the Total Volume to the summary table
                    Range("L" & Summary_Table_Row_Spacer).Value = Total_Volume
                    'Increment the Summary Table Row Spacer so it is ready to print the next occurence of a new
                    'Ticker Name to the summary table
                    Summary_Table_Row_Spacer = Summary_Table_Row_Spacer + 1
                    'Resetting the Total Volume for the next Total Volume calculation
                    Total_Volume = 0
                    'In the case that the Ticker Name of the current row matches the next row
                Else
                    'The Total Volume of the current row is added to the Total Volume being calculated for that Ticker
                    Total_Volume = Total_Volume + Cells(i, 7).Value
                End If
            Next i
        Next ws
    Call Yearly_And_Percentage_Change_Calc
End Sub

Sub Yearly_And_Percentage_Change_Calc()
    'Setting the variable to hold each worksheet in the workbook
    Dim ws As Worksheet
    'Looping through all worksheets
    For Each ws In Worksheets
        'Setting the sheet that the outer loop is currently working on as the active sheet
        ws.Activate
        'Setting variables to hold the Lastrow, Yearly Change, Percentage Change, and space the Yearly and
        'Percentage changes for each Ticker in the summary table.
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Year_Open As Double
        Year_Open = 0
        Dim Year_Close As Double
        Year_Close = 0
        Dim Percentage_Change As Double
        Percentage_Change = 0
        Dim Summary_Table_Row_Spacer As Integer
        Summary_Table_Row_Spacer = 2
        'Printing Yearly Change and Percentage Change headers to their respective columns
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        'Making the Yearly Change and Percentage Change hearder fonts bold
        Range("J1").Font.Bold = True
        Range("K1").Font.Bold = True
        'Auto fitting the column with of the Yearly Change and Percentage Change columns
        Range("J1").EntireColumn.AutoFit
        Range("K1").EntireColumn.AutoFit
            'Looping through each row of the WS that the outer loop is currently working on
            For i = 2 To Lastrow
                'Check if the current row that the inner loop is working on does not match the next row
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'Setting Year_Close to closing price of that stock ticker for that year
                    Year_Close = Cells(i, 6).Value
                    'Calculating the Yearly Change
                    Yearly_Change = Year_Close - Year_Open
                    'Calculating the Percentage Change
                    Percentage_Change = ((Year_Close - Year_Open) / Year_Open) * 100
                    'Print the Yearly Change to the summary table
                    Range("J" & Summary_Table_Row_Spacer).Value = Yearly_Change
                        'Checking the sign of Yearly Change and assigning the corresponding interior color
                        If Yearly_Change > 0 Then
                            Range("J" & Summary_Table_Row_Spacer).Interior.ColorIndex = 4
                        Else
                            Range("J" & Summary_Table_Row_Spacer).Interior.ColorIndex = 3
                        End If
                    'Print the Percentage Change to the summary table
                    Range("K" & Summary_Table_Row_Spacer).Value = Percentage_Change
                    'Increment the Summary Table Row Spacer so it is ready to print the next occurence of a new
                    'Yearly and Percentage changes
                    Summary_Table_Row_Spacer = Summary_Table_Row_Spacer + 1
                    'In the case that the Yearly Change Name of the current row matches the next row
                ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                    Year_Open = Cells(i, 3).Value
                ElseIf Year_Open = 0 And Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                    Year_Open = Cells(i, 3).Value
                End If
            Next i
        Next ws
    Call Greatest_Calc
End Sub

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