Attribute VB_Name = "Module1"
Sub Stock_Homework()


Dim last_row As Long
Dim column As Long
Dim closing_price As Double
Dim opening_price As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim increment As Long

Range("I1") = ("Ticker")
Range("J2") = ("Yearly Change")
Range("K3") = ("Percent Change")
Range("L4") = ("Total Stock volume")

increment = o
column = 1

last_row = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To last_row
        'ticker row
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            increment = increment + 1
            'ticker symbol
            Cells(1 + increment, 9) = Cells(i, 1).Value
            'Difference between end year closing price and the beginning of the year open price
            closing_price = Cells(i, 6).Value
            opening_price = Cells(i - 250, 3).Value
            yearly_change = closing_price - opening_price
            Cells(1 + increment, 10).Value = yearly_change
            'Percentage change from the opening price at the beginning to the closing price at the end of that year.
            percentage_change = (yearly_change / opening_price) * 100
            Cells(1 + increment, 11).Value = percentage_change
            'The total stock volume
            Cells(1 + increment, 12).Formula = "=SUM(" & Range(Cells(i - 250, 7), Cells(i, 7)).Address(False, False) & ")"
            If yearly_change > 0 Then
            'Positive change to Green
            Cells(1 + increment, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
            'Negative change to Red
            Cells(1 + increment, 10).Interior.ColorIndex = 3
                
            End If
            Else
        End If
    Next i
    

Range("R2") = ("Greatest % increase")
Range("R3") = ("Greatest % Decrease")
Range("R4") = ("Greatest total volume")
Range("S1") = ("Ticker")
Range("T1") = ("Value")

End Sub


