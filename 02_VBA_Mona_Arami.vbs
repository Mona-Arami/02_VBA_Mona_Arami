
Sub Level_Hard()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' find the Last Row_Number
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add title header for summary table
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        'declare Variables to hold Values
        Dim Ticker_Name As String
        Dim Volume As Double
        Volume = 0
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double

        Dim Row_Number As Double
        Row_Number = 2

        Dim Column_Number As Integer
        Column_Number = 1

        
        

        'get Open Price value
        Open_Price = Cells(2, Column_Number + 2).Value
         
        Dim i As Long
        For i = 2 To LastRow
         
            If Cells(i + 1, Column_Number).Value <> Cells(i, Column_Number).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, Column_Number).Value
                Cells(Row_Number, Column_Number + 8).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(i, Column_Number + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row_Number, Column_Number + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row_Number, Column_Number + 10).Value = Percent_Change
                    Cells(Row_Number, Column_Number + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volumn
                Volume = Volume + Cells(i, Column_Number + 6).Value
                Cells(Row_Number, Column_Number + 11).Value = Volume
                ' Add one to the summary table row
                Row_Number = Row_Number + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, Column_Number + 2)
                ' reset the Volumn Total
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column_Number + 6).Value
            End If
        Next i
        
        ' find the Last Row_Number of Yearly Change Column for each WS
        LastRow_Sumtable = WS.Cells(Rows.Count, Column_Number + 8).End(xlUp).Row
        ' change the Cell Colors
        For j = 2 To LastRow_Sumtable
            If (Cells(j, Column_Number + 9).Value > 0 Or Cells(j, Column_Number + 9).Value = 0) Then
                Cells(j, Column_Number + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column_Number + 9).Value < 0 Then
                Cells(j, Column_Number + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' set row title for Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column_Number + 14).Value = "Greatest % Increase"
        Cells(3, Column_Number + 14).Value = "Greatest % Decrease"
        Cells(4, Column_Number + 14).Value = "Greatest Total Volume"
        ' set title header for Ticker and Value
        Cells(1, Column_Number + 15).Value = "Ticker"
        Cells(1, Column_Number + 16).Value = "Value"
        ' Look through each rows to find the greatest value 
        For Y = 2 To LastRow_Sumtable
            If Cells(Y, Column_Number + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & LastRow_Sumtable)) Then
                
                Cells(2, Column_Number + 15).Value = Cells(Y, Column_Number + 8).Value
                Cells(2, Column_Number + 16).Value = Cells(Y, Column_Number + 10).Value
                Cells(2, Column_Number + 16).NumberFormat = "0.00%"
            ElseIf Cells(Y, Column_Number + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LastRow_Sumtable)) Then
                
                Cells(3, Column_Number + 15).Value = Cells(Y, Column_Number + 8).Value
                Cells(3, Column_Number + 16).Value = Cells(Y, Column_Number + 10).Value
                
                Cells(3, Column_Number + 16).NumberFormat = "0.00%"
            ElseIf Cells(Y, Column_Number + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRow_Sumtable)) Then
                Cells(4, Column_Number + 15).Value = Cells(Y, Column_Number + 8).Value
                Cells(4, Column_Number + 16).Value = Cells(Y, Column_Number + 11).Value
            End If
        Next Y
        
    Next WS
        
End Sub