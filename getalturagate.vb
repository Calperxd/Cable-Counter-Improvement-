'definir altura do gate no aba Tabela-Gate
'função para preencher altura
Private Sub getaltura()
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("G1").Value = "Altura Nominal"
    For i = 1 To (Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("A1").End(xlDown).Row - 1)
        'altura nominal setada na célula A1 , mudar posteriormente de acordo com o layout
        Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("G" & (i + 1)).Value = Workbooks("taxadeocupacao.xlsm").Sheets("Instruções").Range("A1").Value
        
    Next i
End Sub

