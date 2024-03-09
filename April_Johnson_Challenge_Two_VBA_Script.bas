Attribute VB_Name = "Module1"
Sub tickers()

Dim Sheets As Integer
Dim i As Long
Dim j As Long
Dim k As Long
Dim NoRow As Long
Dim TickerName As Integer
Dim TotalStockVolume As Double
Dim BegYearPrice As Double
Dim EndYearPrice As Double
Dim GreatestInc As Double
Dim GretestDec As Double
Dim GreatestVol As Double
Dim GreatestTicker As String
Dim LeastTicker As String
Dim GreatestVolTicker As String

Sheets = Worksheets.Count

For i = 1 To Sheets

    TickerName = 2
    NoRow = Range("A1").End(xlDown).Row
    TotalStockVolume = 0
    BegYearPrice = 0
    EndYearPrice = 0
    GreatestInc = 0
    GreatestDec = 0
    GreatestVol = 0

    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    For j = 2 To NoRow + 1

        If j = 2 Then
            Cells(TickerName, 9) = Cells(j, 1)
            TotalStockVolume = Cells(j, 7)
            TickerName = TickerName + 1
            BegYearPrice = Cells(j, 3)
        ElseIf Cells(j, 1) <> Cells(j - 1, 1) Then
            Cells(TickerName, 9) = Cells(j, 1)
            Cells(TickerName - 1, 12) = TotalStockVolume
            Range("L" & TickerName - 1).NumberFormat = "0"
            TotalStockVolume = Cells(j, 7)
            EndYearPrice = Cells(j - 1, 6)
            Cells(TickerName - 1, 10) = EndYearPrice - BegYearPrice
            Range("J" & TickerName - 1).NumberFormat = "0.00"
            Cells(TickerName - 1, 11) = (EndYearPrice - BegYearPrice) / BegYearPrice
            BegYearPrice = Cells(j, 3)
            TickerName = TickerName + 1
        Else
            TotalStockVolume = TotalStockVolume + Cells(j, 7)
        End If
               
    
    Next j
  
    For k = 2 To TickerName

        If Cells(k, 10) > 0 Then
            Cells(k, 10).Interior.ColorIndex = 4
        ElseIf Cells(k, 10) < 0 Then
            Cells(k, 10).Interior.ColorIndex = 3
        End If


        If Cells(k, 11) > 0 And Cells(k, 11) > GreatestInc Then
            GreatestInc = Cells(k, 11)
            GreatestTicker = Cells(k, 9)
        End If

        If Cells(k, 11) < 0 And Cells(k, 11) < GreatestDec Then
            GreatestDec = Cells(k, 11)
            LeastTicker = Cells(k, 9)
        End If


        If Cells(k, 12) > GreatestVol Then
            GreatestVol = Cells(k, 12)
            GreatestVolTicker = Cells(k, 9)
        End If

    Cells(k, 11).NumberFormat = "0.00%"


    Next k
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P2") = GreatestTicker
    Range("Q2") = GreatestInc
    Range("P3") = LeastTicker
    Range("Q3") = GreatestDec
    Range("P4") = GreatestVolTicker
    Range("Q4") = GreatestVol
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0"
    
    Columns("J:L").EntireColumn.AutoFit
    Columns("O:Q").EntireColumn.AutoFit
    
    
    If ActiveSheet.Index < Sheets Then
        ActiveSheet.Next.Select
    End If

Next i


End Sub

