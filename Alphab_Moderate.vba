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