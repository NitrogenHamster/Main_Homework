Sub Ticker_Sum()
  
  Dim Ticker As String
  Dim Vol_Total, End_of_Section, PercentChange, YearlyChange As Double
  Dim Summary_Table_Row As Integer
  
  Vol_Total = 0
  Summary_Table_Row = 2
  End_of_Section = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For x = 1 To (Worksheets.Count)
  
    Summary_Table_Row = 2
    
    Worksheets(x).Activate
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Stock Volume"
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
          
            Vol_Total = Vol_Total + Cells(i, 7).Value
            
            YearlyChange = Cells(End_of_Section, 3).Value - Cells(i, 6).Value
               
            Range("I" & Summary_Table_Row).Value = Ticker
            
            If YearlyChange >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            Range("J" & Summary_Table_Row).Value = YearlyChange
            
            If Cells(i, 6).Value = 0 Then
                Range("K" & Summary_Table_Row).Value = "N\A"
            Else
                Range("K" & Summary_Table_Row).Value = Round(((Cells(End_of_Section, 3).Value - Cells(i, 6).Value) / Abs(Cells(i, 6).Value)) * 100, 2)
            End If
            
            Range("L" & Summary_Table_Row).Value = Vol_Total
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Vol_Total = 0
            
            End_of_Section = (i + 1)
           
        Else
          Vol_Total = Vol_Total + Cells(i, 7).Value
        End If
        
    Next
    
  Next
  
End Sub

