Sub stock_data():

    Dim i As Double
    Dim j As Double
    Dim total As Double
    Dim start As Double
    Dim rowCount As Double

    Dim change As Double
    Dim percent As Double

    Dim percentRange As Range
    Dim totalRange As Range
    Dim percentMax As Double
    Dim percentMin As Double
    Dim totalMax As Double

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    j = 0
    total = 0
    start = 2
    change = 0

    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            total = total + Cells(i, 7).Value
            If total = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = 0 & "%"
                Range("L" & 2 + j).Value = 0
                
            Else
                If Cells(start, 3) = 0 Then
                    For Find = start To i
                        If Cells(Find, 3).Value <> 0 Then
                            start = Find
                            Exit For
                        End If
                     Next Find
                End If

                change = (Cells(i, 6) - Cells(start, 3))
                percent = Round((change / Cells(start, 3) * 100), 2)
                start = i + 1

                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(change, 2)
                Range("K" & 2 + j).Value = percent & "%"
                Range("L" & 2 + j).Value = total

                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                End Select

            End If

            j = j + 1
            total = 0
            change = 0

        End If

    Next i

    Set percentRange = Range("K2:K" & Rows.Count)
    percentMax = WorksheetFunction.Max(percentRange)
    percentMin = WorksheetFunction.Min(percentRange)

    Range("Q2").Value = percentMax * 100 & "%"
    Range("Q3").Value = percentMin * 100 & "%"

    Set totalRange = Range("L2:L" & Rows.Count)
    totalMax = WorksheetFunction.Max(totalRange)

    Range("Q4").Value = totalMax

    maxTicker = WorksheetFunction.Match(percentMax, percentRange, 0)
    minTicker = WorksheetFunction.Match(percentMin, percentRange, 0)
    totalTicker = WorksheetFunction.Match(totalMax, totalRange, 0)

    Range("P2") = Cells(maxTicker + 1, 9)
    Range("P3") = Cells(minTicker + 1, 9)
    Range("P4") = Cells(totalTicker + 1, 9)

End Sub