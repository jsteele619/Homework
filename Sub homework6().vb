Sub homework6()

Dim Ticker As String
Dim Change As Integer
Dim Percent As Single
Dim Total_Stock As Double
Dim Closing As Single
Dim Opening As Single
Dim Sum As Integer
Dim Lastrow As Long

Total_Stock = 0
Sum = 2

For Each Sheet In Worksheets

Lastrow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row

For j = 2 To Lastrow

    If Sheet.Cells(j - 1, 1).Value <> Sheet.Cells(j, 1) Then
        Opening = Sheet.Cells(j, 3)
    End If

    If Sheet.Cells(j + 1, 1).Value <> Sheet.Cells(j, 1).Value Then
        Ticker = Sheet.Cells(j, 1)
        Sheet.Cells(Sum, 9) = Ticker
        Closing = Sheet.Cells(j, 6).Value
        Sheet.Cells(Sum, 10).Value = Closing - Opening
        Sheet.Cells(Sum, 11).Value = Sheet.Cells(Sum, 10).Value / Opening
    
    If Sheet.Cells(Sum, 10).Value > 0 Then
        Sheet.Cells(Sum, 10).Interior.ColorIndex = 4
    End If
        
     If Sheet.Cells(Sum, 10).Value < 0 Then
        Sheet.Cells(Sum, 10).Interior.ColorIndex = 3
    End If
    
        Total_Stock = Total_Stock + Sheet.Cells(j, 7).Value
        Sheet.Cells(Sum, 12) = Total_Stock
        Total_Stock = 0
        Sum = Sum + 1
     
    Else: Sheet.Cells(j + 1, 1).Value = Sheet.Cells(j, 1).Value
        Total_Stock = Total_Stock + Sheet.Cells(j, 7).Value
    End If
    
    Sheet.Cells(Sum, 10).NumberFormat = "0.00#"
    Sheet.Cells(Sum, 11).NumberFormat = "0.00%"
Next j

    Sum = 2

    Sheet.Cells(1, 9).Value = "Ticker"
    Sheet.Cells(1, 10).Value = "Yearly Change"
    Sheet.Cells(1, 11).Value = "Percent Change"
    Sheet.Cells(1, 12).Value = "Total Stock Volume"

Next Sheet
End Sub