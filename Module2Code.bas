Attribute VB_Name = "Module1"
Sub analysis():

For Each w In Worksheets
    Dim vol_stock As Double
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim SummaryTableRow As Integer
    Dim WorksheetName As String
    Dim GI As Double
    Dim GIticker As String
    Dim GD As Double
    Dim GDticker As String
    Dim GTV As Double
    Dim GTVticker As String
    w.Range("K:K").NumberFormat = "0.00%"
    w.Range("Q2:Q3").NumberFormat = "0.00%"
    w.Range("I1").Value = "Ticker"
    w.Range("J1").Value = "Yearly Change"
    w.Range("K1").Value = "Percent Change"
    w.Range("L1").Value = "Total Stock Volume"
    w.Range("P1").Value = "Ticker"
    w.Range("Q1").Value = "Value"
    w.Range("O2").Value = "Greatest % Increase"
    w.Range("O3").Value = "Greatest % Decrease"
    w.Range("O4").Value = "Greatest Total Volume"
    LastRow = w.Cells(Rows.Count, 1).End(xlUp).Row
    SummaryTableRow = 2
    GI = 0
    GD = 0
    GTV = 0
    openprice = w.Cells(2, 3).Value

    'ticker and total stock volume
    For i = 2 To LastRow
        If w.Cells(i + 1, 1).Value <> w.Cells(i, 1).Value Then
            vol_stock = vol_stock + w.Cells(i, 7).Value
            ticker = w.Cells(i, 1).Value
            closeprice = w.Cells(i, 6).Value
            w.Range("I" & SummaryTableRow).Value = ticker
            w.Range("L" & SummaryTableRow).Value = vol_stock
            yearlychange = closeprice - openprice
            percentchange = ((closeprice - openprice) / openprice)
            w.Range("J" & SummaryTableRow).Value = yearlychange
            w.Range("K" & SummaryTableRow).Value = percentchange
            'important to note that the calculation is done before the new openprice is taken
            openprice = w.Cells(i + 1, 3).Value
            SummaryTableRow = SummaryTableRow + 1
            vol_stock = 0
        Else
            vol_stock = vol_stock + w.Cells(i, 7).Value
        End If
    Next i
    
    For i = 2 To SummaryTableRow - 1
        'conditional formatting of yearly change
        If w.Cells(i, 10).Value >= 0 Then
            w.Cells(i, 10).Interior.ColorIndex = 4
        Else
            w.Cells(i, 10).Interior.ColorIndex = 3
        End If
   
        'greatest % increase/decrease
        If w.Cells(i, 11).Value > GI Then
            GI = w.Cells(i, 11).Value
            GIticker = w.Cells(i, 9).Value
        ElseIf w.Cells(i, 11).Value < GD Then
            GD = w.Cells(i, 11).Value
            GDticker = w.Cells(i, 9).Value
        End If
        
        'greatest total volume
        If w.Cells(i, 12).Value > GTV Then
            GTV = w.Cells(i, 12).Value
            GTVticker = w.Cells(i, 9).Value
        End If
    Next i
    
    w.Cells(2, 16).Value = GIticker
    w.Cells(2, 17).Value = GI
    w.Cells(3, 16).Value = GDticker
    w.Cells(3, 17).Value = GD
    w.Cells(4, 16).Value = GTVticker
    w.Cells(4, 17).Value = GTV
    
    w.Columns("A:Q").AutoFit
    
Next w

MsgBox "fin"

End Sub

