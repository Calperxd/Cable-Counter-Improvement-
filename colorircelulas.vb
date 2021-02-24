'esta função colore as células qaue estão definidas como critério

Private Sub colorir()
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B2:B" & Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("B1").End(xlDown).Row).Interior.Color = vbWhite
    For i = 1 To Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("B2").End(xlDown).Row - 1
        If Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("I" & (i + 1)).Value > Workbooks("taxadeocupacao.xlsm").Sheets("Instruções").Range("A2").Value Then
            Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("I" & (i + 1)).Interior.Color = vbRed
        End If
    Next i
End Sub

