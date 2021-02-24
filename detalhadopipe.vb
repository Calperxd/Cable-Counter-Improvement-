Public pipes() As String 'holds all the gates
Dim c As Range

Private Sub fillArrayOfPipes(pipeClicked As String)
    Dim primeiroEndereco As String
    Dim c As Range
    Dim index As Integer
    index = 0               'variável pra controlar a quantidade de gates achados
    With Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("A:A")
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("A:A").Find(pipeClicked, LookAt:=xlWhole)
        If Not c Is Nothing Then
            primeiroEndereco = c.Address
            ReDim Preserve pipes(index)
            pipes(0) = c.Cells.Offset(0, 1).Value
            Do
                Set c = .FindNext(c)
                If c.Address <> primeiroEndereco Then
                    index = index + 1   'achei o primeiro
                    ReDim Preserve pipes(index)
                    pipes(UBound(pipes)) = c.Cells.Offset(0, 1)
                End If
            Loop While Not c Is Nothing And c.Address <> primeiroEndereco
        End If
    End With
End Sub
'criar 11 funções para preencher cada campo do relatório de leitos
'cada função deve fazer apenas um operação
'os leitos estão armazenados no array pipes
'trocar array de string para array variant para usar a função .hasElement


Private Sub gate(inputGate As String)
    c.Cells.Offset(0, 1) = inputGate
End Sub

Private Sub nivel(inputGate As String)
    Dim lv As String
    lv = Split(inputGate, "/")(2) 'Recebe apenas o "LV2.4" /P101.1/LV2.4
    c.Cells.Offset(1, 0) = lv
End Sub
Private Sub leito(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = d.Cells.Offset(0, 2)
End Sub
Private Sub areaUtil(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = (1 - (d.Cells.Offset(0, 7).Value / 100)) * d.Cells.Offset(0, 6)
End Sub
Private Sub areaOcupada(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = (d.Cells.Offset(0, 7).Value / 100) * d.Cells.Offset(0, 6)
End Sub
Private Sub ocupacao(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = d.Cells.Offset(0, 7)
End Sub
Private Sub criterio()
    c.Cells.Offset(1, 0) = Workbooks("taxadeocupacao.xlsm").Sheets("Instruções").Range("A2").Value
End Sub
Private Sub comprimento(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = d.Cells.Offset(0, 3)
End Sub
Private Sub pesoLeitoKgm(inputGate As String)
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Tabela-Gate").Range("A:A").Find(Split(d.Cells.Offset(0, 2), "mm")(0), LookAt:=xlWhole)
    c.Cells.Offset(1, 0) = d.Cells.Offset(0, 5)

End Sub

Private Sub pesoCabosKgm(inputGate As String)
    Dim d As Range
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Select
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range(d.Address).Select
    lastLine = Workbooks("taxadeocupacao.xlsm").Sheets(Sheets.Count).Range("D2").End(xlDown).Row
    c.Cells.Offset(1, 0) = (WorksheetFunction.Sum(Workbooks("taxadeocupacao.xlsm").Sheets(Sheets.Count).Range("D2:D" & lastLine))) / d.Cells.Offset(0, 3)
    Application.DisplayAlerts = False
    Workbooks("taxadeocupacao.xlsm").Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub

Private Sub pesoTotalKgm()
    c.Cells.Offset(1, 0) = c.Cells.Offset(1, -2) + c.Cells.Offset(1, -1)

End Sub


Private Sub fillTheAreas()
    Dim j As Long
    j = 5
    For i = 0 To UBound(pipes)
        
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("GATE", LookAt:=xlWhole)
        gate (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Nível", LookAt:=xlWhole)
        nivel (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Leito", LookAt:=xlWhole)
        leito (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Área útil mm²", LookAt:=xlWhole)
        areaUtil (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Área ocupada mm²", LookAt:=xlWhole)
        areaOcupada (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("% Ocupação", LookAt:=xlWhole)
        ocupacao (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("% Critério", LookAt:=xlWhole)
        criterio
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Comprimento m", LookAt:=xlWhole)
        comprimento (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Peso Leito kg/m", LookAt:=xlWhole)
        pesoLeitoKgm (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Peso cabos kg/m", LookAt:=xlWhole)
        pesoCabosKgm (pipes(i))
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Peso Total kg/m", LookAt:=xlWhole)
        pesoTotalKgm
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("Camadas", LookAt:=xlWhole)
        colorirRelatorioGates
        Workbooks("taxadeocupacao.xlsm").Worksheets("Rascunho").Range("A5:G15").Copy _
            Destination:=Workbooks("taxadeocupacao.xlsm").Worksheets("Gates-Resumo").Range("A" & j)
        j = j + 11

    Next i
    
End Sub
Private Sub colorirRelatorioGates()
    Dim d As Range
    Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("E:E").Find("% Ocupação", LookAt:=xlWhole)
    Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("A:G").Find("% Critério", LookAt:=xlWhole)
    Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("E11").Interior.Color = vbWhite
    
    temp1 = d.Cells.Offset(1, 0).Value
    temp2 = c.Cells.Offset(1, 0).Value
    If temp1 > temp2 Then
        Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("E11").Interior.Color = vbRed
    Else
        Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("E11").Interior.Color = vbWhite
    End If
End Sub


Private Sub relatorio_detalhado_pipe(strAdr As String, pipeClicked As String)
    
    teste = Workbooks("taxadeocupacao.xlsm").Sheets("Gates-Resumo").Cells.SpecialCells(xlCellTypeLastCell).Address
    Workbooks("taxadeocupacao.xlsm").Sheets("Gates-Resumo").Range("A5:" & teste).Delete
    fillArrayOfPipes (pipeClicked)
    Workbooks("taxadeocupacao.xlsm").Sheets("Rascunho").Range("E11").Interior.Color = vbWhite
    If Not pipes(0) = "" Then
        fillTheAreas
    End If
    ReDim pipes(0)
End Sub

