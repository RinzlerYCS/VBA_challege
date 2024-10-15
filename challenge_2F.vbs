Sub RunForAllSheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        Call challenge_2
    Next ws

End Sub

Sub challenge_2()

    Dim LastRow As Long
    Dim vol1 As LongLong
    Dim counter As Integer
    Dim ticker As String
    Dim columnn As Integer
    Dim open_value As Double
    Dim close_value As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    
    
    Range("J2:M2").ClearContents
    Range("P2:Q5").ClearContents
   
    Range("J2").Value = "Ticker"
    Range("K2").Value = "Quarterly change"
    Range("L2").Value = "Percent change"
    Range("M2").Value = "Total stock volume"
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    counter = 0
    columnn = 2
   
    For j = 2 To LastRow
        ticker = Range("A" & j).Value
        vol1 = (Range("G" & j).Value) + vol1
        counter = counter + 1

        If counter = 1 Then
            open_value = Range("C" & j).Value
        End If
        
        If counter = 62 Then
        
            close_value = Range("F" & j).Value
            
            columnn = columnn + 1
            Range("J" & columnn).Value = Range("A" & j).Value
            Range("K" & columnn).Value = (close_value - open_value)
            Range("L" & columnn).Value = ((close_value - open_value) / open_value)
            Range("L" & columnn).NumberFormat = "0.00%"
            Range("M" & columnn).Value = vol1
            
            If (close_value - open_value) < 0 Then
                Range("K" & columnn).Interior.Color = RGB(255, 88, 51)
            ElseIf Range("K" & columnn) > 0 Then
                Range("K" & columnn).Interior.Color = RGB(60, 255, 51)
            End If

            counter = 0
            vol1 = 0
            open_value = 0
            close_value = 0
        End If
    Next j
    Columns("M:M").EntireColumn.AutoFit

    Dim greatest_increase_TICKER As String
    Dim greatest_increase_VALUE As Double
    Dim greatest_decrease_TICKER As String
    Dim greatest_decrease_VALUE As Double
    Dim greatest_total_volume_TICKER As String
    Dim greatest_total_volume_Value As LongLong
    
    greatest_total_volume_Value = 0
    greatest_decrease_VALUE = 0
    greatest_increase_VALUE = 0
    
    For i = 3 To LastRow
        If (Range("L" & i).Value) > greatest_increase_VALUE Then
            greatest_increase_VALUE = Range("L" & i).Value
            greatest_increase_TICKER = Range("J" & i).Value
        End If
        
        If (Range("L" & i).Value) < greatest_decrease_VALUE Then
            greatest_decrease_VALUE = Range("L" & i).Value
            greatest_decrease_TICKER = Range("J" & i).Value
        End If
        
        If (Range("M" & i).Value) > greatest_total_volume_Value Then
            greatest_total_volume_Value = Range("M" & i).Value
            greatest_total_volume_TICKER = Range("J" & i).Value
        End If
    Next i
            
    Range("P2").Value = "Ticker"
    Range("Q2").Value = "Value"
    
    Range("O3").Value = "Greatest % increase"
    Range("P3").Value = greatest_increase_TICKER
    Range("Q3").Value = greatest_increase_VALUE
    
    Range("O4").Value = "Greatest % decrease"
    Range("P4").Value = greatest_decrease_TICKER
    Range("Q4").Value = greatest_decrease_VALUE
    
    Range("O5").Value = "Greatest total volume"
    Range("P5").Value = greatest_total_volume_TICKER
    Range("Q5").Value = greatest_total_volume_Value
    
    Range("Q3:Q5").NumberFormat = "0.00%"
    
End Sub
