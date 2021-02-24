'definir area util do gate no aba Tabela-Gate
'função para preencher altura
Private Sub getareautil()
    Dim c As Range      'variável que armazena o resultado da procura na planilha Tabela-GAte
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("H1").Value = "Área útil mm²"
    For i = 1 To (Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("A1").End(xlDown).Row - 1)
        'Área útil mm² setada na célula A2 , mudar posteriormente de acordo com o layout
        hbore = Split(Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("C" & (i + 1)).Value, ".00")(0)
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Tabela-Gate").Range("A:A").Find(hbore)
        Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("H" & (i + 1)).Value = c.Cells.Offset(0, 4).Value
    Next i
End Sub
