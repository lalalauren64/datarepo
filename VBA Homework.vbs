Sub Multiple_YR_Stock_Data()

    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim last_row As Long
    
    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To last_row
          
     'For i = 2 To 17800
     
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            Ticker = Cells(i, 1).Value

            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            Range("I" & Summary_Table_Row).Value = Ticker

            Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

            Summary_Table_Row = Summary_Table_Row + 1

            Total_Stock_Volume = 0

        Else
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If

    Next i

End Sub

Sub multi_sheet()

End Sub

   Dim last_row As Long
        
        For Each ws In Worksheets
               
            last_row = Cells(Rows.Count, "A").End(xlUp).Row
        
            MsgBox (last_row)

        Next ws
   
End Sub
