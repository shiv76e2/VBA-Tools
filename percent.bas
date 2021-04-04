Sub percent()
    Dim MaxRow As Integer
    MaxRow = Cells(Rows.Count, 2).End(xlUp).Row
    Dim MaxCol As Integer
    Dim mol As Integer
    Dim den As Integer
    Dim TotMol As Integer: TotMol = 0
    Dim TolDen As Integer: TotDen = 0
    
    
    Dim i As Integer
    Dim j As Integer
    For i = 3 To MaxRow
        MaxCol = Cells(i, 2).End(xlToRight).Column
        mol = 0
        den = 0
        For j = 3 To MaxCol
            den = den + 1
            TotDen = TotDen + 1
            If Cells(i, j).Interior.ColorIndex <> xlColorIndexNone Then
                mol = mol + 1
                TotMol = TotMol + 1
            End If
        Next
        Cells(i, 1) = mol / den
    Next
    Cells(17, 8) = TotMol / TotDen '総合の進捗割合を表示する
End Sub