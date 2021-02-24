Dim tempcabos() As String 'armazena os cabos em unidades pro ex: 2x(2x45) + 1x(2x100)
                              'tempcabos(0) = 2x(2x45) primeiro cabo
                              'tempcabos(1) = 1x(2x100) segundo cabo
Dim qtdcabo() As Integer 'armazena as quantidades de cabos pro ex: 2x(2x45) + 1x(2x100)
                              'tempcabos(0) = 2 - primeiro cabo
                              'tempcabos(1) = 1 - segundo cabo
Dim area As Double
Public sheet As Worksheet


Private Sub preencher_cabos(str As String) 'recebe o gate como parâmetro, e executa uma busca
    Dim c As Range
    Dim primeiroEndereco As String
    Dim gate As String
    Dim linha As Integer
    
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    sheet.Name = Replace(Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range(str).Value, "/", "")
    temp = sheet.Name           'esta variável recebe o nome da planilha criada
    sheet.Range("A1").Value = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range(str).Value
    sheet.Range("B1").Value = "Área ocupada [mm²]"
    sheet.Range("C1").Value = "Taxa de Ocupação [%]"
    sheet.Range("D1").Value = "Peso do Cabo [kg]"
    'pega o nome do gate clickado
    gate = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range(str).Value
    gate = Replace(gate, "/", "", , 1)
    linha = i 'pega a linha do gate
    With Workbooks("taxadeocupacao.xlsm").Sheets("Cabo-Rota").Range("A:XFD")
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Cabo-Rota").Range("A:XFD").Find(gate, LookAt:=xlWhole) 'a variável gate entra como parametro para essa função para procurar na tabela de Cabo-Rotas
            'se a variável C não retornar nothing quer dizer que ela achou cabos
            'armazenar primeiro endereço e depois ir para o próximo endereço achado
        If Not c Is Nothing Then
            primeiroEndereco = c.Address
            ReDim Preserve arraydecabos(qtddecabos)
            arraydecabos(0) = c.Cells.Offset(0, (c.Column * -1) + 1).Value
            Do
                'chamar função para cálcular cabos de outro módulo
                Set c = .FindNext(c)
                qtddecabos = UBound(arraydecabos)
                If c.Address <> primeiroEndereco Then
                    ReDim Preserve arraydecabos(qtddecabos + 1)
                    qtddecabos = qtddecabos + 1
                    arraydecabos(qtddecabos) = c.Cells.Offset(0, (c.Column * -1) + 1).Value
                End If
                Loop While Not c Is Nothing And c.Address <> primeiroEndereco
                    For i = 0 To UBound(arraydecabos)
                        Workbooks("taxadeocupacao.xlsm").Sheets(temp).Range("A" & (i + 2)).Value = arraydecabos(i)
                    Next i
                qtddecabos = 0
            End If
        End With
    sheet.Range("A:XFD").Columns.AutoFit
End Sub

Private Sub acumuladordeareaatual()
    Dim c As Range
    Dim temp As String

    ReDim Preserve qtdcabo(UBound(tempcabos))
    
    If StrComp(tempcabos(UBound(tempcabos)), "sh)", vbBinaryCompare) = 0 Then
        tempcabos(0) = tempcabos(0) & "+"
        tammax = (UBound(tempcabos) - 1)
    Else
        tammax = UBound(tempcabos)
    End If
    For i = 0 To tammax
        temp = Split(tempcabos(i), "x(")(1)
        temp = Replace(temp, ")", "") 'recebe apenas o 2x45)
        qtdcabo(i) = CInt(Split(tempcabos(i), "x(")(0)) 'recebe apenas a quantidade de cabo de cada dimensão
        tempcabos(i) = Replace(Split(tempcabos(i), "x(")(1), ")", "")
    
    Next i
    For j = 0 To tammax
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Tabela-Cabo").Range("C:C").Find(tempcabos(j), LookAt:=xlPart)
        area = CDbl(c.Cells.Offset(0, 2).Value) + area
        area = area * qtdcabo(j)
    Next j

End Sub


'essa função calcula apenas a ocupação do array atual
Private Sub calcular_parametros()
    Dim dimensaodocabo As String
    For i = 0 To UBound(arraydecabos)
        dimensaodocabo = Split(arraydecabos(i), "_")(2)
        tempcabos = Split(dimensaodocabo, "+")
        acumuladordeareaatual
        Workbooks("taxadeocupacao.xlsm").Sheets(sheet.Name).Range("B" & (i + 2)).Value = area
        area = 0
    Next i
    ReDim qtdcabo(0)
End Sub
Private Sub taxadeocup()
    Dim c As Range
    procurar = sheet.Range("A1").Value
    Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(procurar, LookAt:=xlWhole)
    For i = 0 To UBound(arraydecabos)
        sheet.Range("C" & (i + 2)).Value = Round((sheet.Range("B" & (i + 2)) / c.Cells.Offset(0, 6)) * 100, 2)
    Next i
End Sub

Private Sub calcula_peso()
    Dim c As Range
    Dim temp As String

    ReDim Preserve qtdcabo(UBound(tempcabos))
    
    If StrComp(tempcabos(UBound(tempcabos)), "sh)", vbBinaryCompare) = 0 Then
        tempcabos(0) = tempcabos(0) & "+"
        tammax = (UBound(tempcabos) - 1)
    Else
        tammax = UBound(tempcabos)
    End If
    For i = 0 To tammax
        temp = Split(tempcabos(i), "x(")(1)
        temp = Replace(temp, ")", "") 'recebe apenas o 2x45)
        qtdcabo(i) = CInt(Split(tempcabos(i), "x(")(0)) 'recebe apenas a quantidade de cabo de cada dimensão
        tempcabos(i) = Replace(Split(tempcabos(i), "x(")(1), ")", "")
    
    Next i
    For j = 0 To tammax
        Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Tabela-Cabo").Range("C:C").Find(tempcabos(j), LookAt:=xlPart)
        area = CDbl(c.Cells.Offset(0, 2).Value) + area
        area = area * qtdcabo(j)
    Next j

End Sub

'essa função calcula o peso do pack do cabo
Private Sub calcular_peso_pack_cabo()
    Dim dimensaodocabo As String
    For i = 0 To UBound(arraydecabos)
        dimensaodocabo = Split(arraydecabos(i), "_")(2)
        tempcabos = Split(dimensaodocabo, "+")
        calcula_peso
        Workbooks("taxadeocupacao.xlsm").Sheets(sheet.Name).Range("B" & (i + 2)).Value = area
        area = 0
    Next i
    ReDim qtdcabo(0)
End Sub
'função para calcular o peso do cabo considerando o tamanho
Private Sub getpeso(inputGate As String) 'recebe o gate clickado como parametro
    Dim tammax As Integer
    Dim Peso As Double             'recebe o peso apenas do primeiro cabo
                                    'peso = 41kg, 2x(2x45) primeiro cabo
                                    'peso = 57kg, 1x(2x100) segundo cabo
    Dim pesototal As Double
    pesototal = 0
    For k = 0 To (Workbooks("taxadeocupacao.xlsm").Sheets(Sheets.Count).Range("A1").End(xlDown).Row - 2)
        'vetor que recebe as dimensões do cabo
        'pos(0) = 1x(4x25)
        'pos(1) = 1x(4x35)
        tempcabos = Split(Split(sheet.Range("A" & (k + 2)).Value, "_")(2), "+")
        ReDim Preserve qtdcabo(UBound(tempcabos))
        If StrComp(tempcabos(UBound(tempcabos)), "sh)", vbBinaryCompare) = 0 Then
            tempcabos(0) = tempcabos(0) & "+"
            tammax = (UBound(tempcabos) - 1)
        Else
                tammax = UBound(tempcabos)
        End If
        For i = 0 To tammax
            temp = Split(tempcabos(i), "x(")(1)
            temp = Replace(temp, ")", "") 'recebe apenas o 2x45)
            qtdcabo(i) = CInt(Split(tempcabos(i), "x(")(0)) 'recebe apenas a quantidade de cabo de cada dimensão
            tempcabos(i) = Replace(Split(tempcabos(i), "x(")(1), ")", "")
        
        Next i
        For j = 0 To tammax
            Set c = Workbooks("taxadeocupacao.xlsm").Sheets("Tabela-Cabo").Range("C:C").Find(tempcabos(j), LookAt:=xlPart)
            Set d = Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("B:B").Find(inputGate, LookAt:=xlWhole)
            Peso = CDbl(c.Cells.Offset(0, 3).Value) * (d.Cells.Offset(0, 3) / 1000) * qtdcabo(j) ' * comprimento do gate em km
            pesototal = pesototal + Peso
            Peso = 0
        Next j
         Workbooks("taxadeocupacao.xlsm").Sheets(Sheets.Count).Range("D" & (k + 2)).Value = Round(pesototal, 2)
    Next k
    
End Sub




Private Sub relatorio_detalhado_gate(str As String, gateclicked As String)         'str é uma váriavel que recebe o endereço do item clickado
    Dim existecabo As Boolean
    
    
    qtddecabos = 0
    preencher_cabos (str)
    
    'verifica se o vetor arraydecabos tem algum cabo armazenado nele
    If Not arraydecabos(0) = "" Then
        calcular_parametros
        taxadeocup
        getpeso (gateclicked)
    End If
    ReDim arraydecabos(0)
    sheet.Range("A:XFD").Columns.AutoFit
End Sub
