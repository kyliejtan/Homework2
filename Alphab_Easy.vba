Sub Total_Volume_Calc()
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
        'Total_Volume = 0
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
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
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
End Sub




