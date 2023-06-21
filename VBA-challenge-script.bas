Attribute VB_Name = "Module1"
Sub StockRunTest()
' Assuming all data is ordered by stock ticker and by date
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Stock Volume"

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Column_Format = 1

Open_Price = Cells(2, 3)
Total_Stock = Cells(2, 7)
Yearly_Change = 0

For i = 2 To LastRow
    If Cells(i + 1, 1) <> Cells(i, 1) Then
        Column_Format = Column_Format + 1
        Range("I" & Column_Format) = Cells(i, 1)
        Yearly_Change = Cells(i, 6) - Open_Price
        Range("J" & Column_Format) = Yearly_Change
        If Yearly_Change > 0 Then
            Range("J" & Column_Format).Interior.ColorIndex = 4
        ElseIf Yearly_Change < 0 Then
            Range("J" & Column_Format).Interior.ColorIndex = 3
        ' 0 change will be left as nothing
        End If
        Range("K" & Column_Format) = Yearly_Change / Open_Price
        Range("K" & Column_Format).NumberFormat = "0.00%"
        Range("L" & Column_Format) = Total_Stock
        Open_Price = Cells(i + 1, 3)
        Total_Stock = Cells(i + 1, 7)
    Else
        Total_Stock = Total_Stock + Cells(i + 1, 7)
    End If
Next i

LastRow_2 = Cells(Rows.Count, 9).End(xlUp).Row


Greatest = Cells(2, 11)
Greatest_Name = Cells(2, 9)
Smallest = Cells(2, 11)
Smallest_Name = Cells(2, 9)
Largest = Cells(2, 12)
Largest_Name = Cells(2, 9)


For i = 2 To LastRow_2
    If Cells(i + 1, 11) > Greatest Then
        Greatest = Cells(i + 1, 11)
        Greatest_Name = Cells(i + 1, 9)
    ElseIf Cells(i + 1, 11) < Smallest Then
        Smallest = Cells(i + 1, 11)
        Smallest_Name = Cells(i + 1, 9)
    End If
    If Cells(i + 1, 12) > Largest Then
        Largest = Cells(i + 1, 12)
        Largest_Name = Cells(i + 1, 9)
    End If
Next i

Range("P1") = "Ticker"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("P2") = Greatest_Name
Range("Q2") = Greatest
Range("Q2").NumberFormat = "0.00%"
Range("O3") = "Greatest % Decrease"
Range("P3") = Smallest_Name
Range("Q3") = Smallest
Range("Q3").NumberFormat = "0.00%"
Range("O4") = "Greatest Total Value"
Range("P4") = Largest_Name
Range("Q4") = FormatCurrency(Largest)

End Sub

Sub YearlyStockRun()
For Each ws In Worksheets
    Dim WorksheetName As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    
    Column_Format = 1

    Open_Price = ws.Cells(2, 3)
    Total_Stock = ws.Cells(2, 7)
    Yearly_Change = 0
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Column_Format = Column_Format + 1
            ws.Range("I" & Column_Format) = ws.Cells(i, 1)
            Yearly_Change = ws.Cells(i, 6) - Open_Price
            ws.Range("J" & Column_Format) = Yearly_Change
            If Yearly_Change > 0 Then
                ws.Range("J" & Column_Format).Interior.ColorIndex = 4
            ElseIf Yearly_Change < 0 Then
                ws.Range("J" & Column_Format).Interior.ColorIndex = 3
            End If
            ws.Range("K" & Column_Format) = Yearly_Change / Open_Price
            ws.Range("K" & Column_Format).NumberFormat = "0.00%"
            ws.Range("L" & Column_Format) = Total_Stock
            Open_Price = ws.Cells(i + 1, 3)
            Total_Stock = ws.Cells(i + 1, 7)
        Else
            Total_Stock = Total_Stock + ws.Cells(i + 1, 7)
        End If
    Next i
    LastRow_2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
    Greatest = ws.Cells(2, 11)
    Greatest_Name = ws.Cells(2, 9)
    Smallest = ws.Cells(2, 11)
    Smallest_Name = ws.Cells(2, 9)
    Largest = ws.Cells(2, 12)
    Largest_Name = ws.Cells(2, 9)
    
    
    For i = 2 To LastRow_2
        If ws.Cells(i + 1, 11) > Greatest Then
            Greatest = ws.Cells(i + 1, 11)
            Greatest_Name = ws.Cells(i + 1, 9)
        ElseIf ws.Cells(i + 1, 11) < Smallest Then
            Smallest = ws.Cells(i + 1, 11)
            Smallest_Name = ws.Cells(i + 1, 9)
        End If
        If ws.Cells(i + 1, 12) > Largest Then
            Largest = ws.Cells(i + 1, 12)
            Largest_Name = ws.Cells(i + 1, 9)
        End If
    Next i
    
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("P2") = Greatest_Name
    ws.Range("Q2") = Greatest
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("P3") = Smallest_Name
    ws.Range("Q3") = Smallest
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("O4") = "Greatest Total Value"
    ws.Range("P4") = Largest_Name
    ws.Range("Q4") = FormatCurrency(Largest)
    
Next ws
End Sub



