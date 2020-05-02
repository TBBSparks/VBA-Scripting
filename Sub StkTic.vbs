Sub StkTic()

' Set CurrentWs as a worksheet object variable.
    Dim CurrentWs As Worksheet

' Loop through all of the worksheets in the active workbook.
    For Each CurrentWs In Worksheets

    'set variables
    DIM Ticker as String
    Ticker = " "
    DIM YearChange as Long
    YearChange = 0
    DIM PercChange as Long
    PercChange = 0
    DIM TotVol as Double
    TotVol = 0
    DIM MaxPercName as String
    MaxPercName = " "
    DIM MinPercName as String
    MinPercName = " "
    DIM MaxVolName as String
    MaxVolName = " "
    DIM MaxPerc as Double
    MaxPerc = 0
    DIM MinPerc as Double
    MinPerc = 0
    DIM MaxVol as Double
    MaxVol = 0
    DIM OpenPrice as Double
    OpenPrice = 0
    DIM ClosePrice as Double
    ClosePrice = 0
    DIM SumTblRow as long
    SumTblRow = 2
    DIM LastRow as Long
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    DIM i as Long

    'Print out headers
    CurrentWs.Range("I1").Value = "Ticker"
    CurrentWs.Range("J1").Value = "Yearly Change"
    CurrentWs.Range("K1").Value = "Percent Change"
    CurrentWs.Range("L1").Value = "Total Stock Volume"
    CurrentWs.Range("O2").Value = "Greatest % Increase"
    CurrentWs.Range("O3").Value = "Greatest % Decrease"
    CurrentWs.Range("O4").Value = "Greatest Total Volume"
    CurrentWs.Range("P1").Value = "Ticker"
    CurrentWs.Range("Q1").Value = "Value"


' Set variable OpenPrice to be first price of first stock
OpenPrice = CurrentWs.Cells(2, 3).Value

' Loop through cell by cell setting closing, calculating close vs open, and summing the total volume
For i = 2 to LastRow
    ' If the next ticker value equals the current value perform math
    If CurrentWs.Cells(i + 1, 1).Value = CurrentWs.Cells(i, 1).Value Then

        Ticker = CurrentWs.Cells(i, 1).Value
        ClosePrice = CurrentWs.Cells(i, 6).Value
        YearChange = ClosePrice - OpenPrice
        TotVol = TotVol + CurrentWs.Cells(i, 7).Value
    ' If the next ticker value is not equal to the current perform math one last time then fill cells
    Else
        ClosePrice = CurrentWs.Cells(i, 6).Value
        YearChange = ClosePrice - OpenPrice
        TotVol = TotVol + CurrentWs.Cells(i, 7).Value
        If (OpenPrice > 0) Then
            PercChange = (YearChange / OpenPrice) * 100
        Else PercChange = 0
        End If
        CurrentWs.Range("I" & SumTblRow).Value = Ticker
        CurrentWs.Range("J" & SumTblRow).Value = YearChange
        ' Color green for stock gain or red for loss
        If (YearChange > 0) Then
            CurrentWs.Range("J" & SumTblRow).Interior.ColorIndex = 4
        ElseIf (YearChange <= 0) Then
            CurrentWs.Range("J" & SumTblRow).Interior.ColorIndex = 3
        End If
        CurrentWs.Range("K" & SumTblRow).Value = (CStr(PercChange) & "%")
        CurrentWs.Range("L" & SumTblRow).Value = TotVol
                
        ' Check to see if the current stock meets max min change conditions and store if so
        If (PercChange > MaxPerc) Then
            MaxPerc = PercChange
            MaxPercName = Ticker
        ElseIf (PercChange < MinPerc) Then
            MinPerc = PercChange
            MinPercName = Ticker
        End If
                       
        If (TotVol > MaxVol) Then
            MaxVol = TotVol
            MaxVolName = Ticker
        End If

        ' Increment the results table by 1 for next print
        SumTblRow = SumTblRow + 1
        ' reset
        YearChange = 0
        ' reset
        ClosePrice = 0
        ' reset
        TotVol = 0
        'reset
        PercChange = 0
        ' Get next open price for next stock which is one cell ahead of current position
        OpenPrice = CurrentWs.Cells(i + 1, 3).Value

    End If

Next i


' put max min values in proper cells on current worksheet
CurrentWs.Range("Q2").Value = (CStr(MaxPerc) & "%")
CurrentWs.Range("Q3").Value = (CStr(MinPerc) & "%")
CurrentWs.Range("P2").Value = MaxPercName
CurrentWs.Range("P3").Value = MinPercName
CurrentWs.Range("Q4").Value = MaxVol
CurrentWs.Range("P4").Value = MaxVolName

        
Next CurrentWs

End Sub