Public arraydecabos() As String
Public criada As Boolean
Public RodarMacroClicked As Boolean
'exclui excesso de planilhas criadas
Private Sub excluirPlanilhas()
    For i = Workbooks("taxadeocupacao.xlsm").Sheets.Count To 9 Step -1
        Application.DisplayAlerts = False
        Workbooks("taxadeocupacao.xlsm").Sheets(i).Delete
        Application.DisplayAlerts = True
    Next i
End Sub
'procurar cabos para preencher o vetor arraydecabos() que é redimensionável de acordo com a quantidade de cabo
Private Sub procurar_cabo_gate()
  Dim c As Range
  Dim primeiroEndereco As String
  Dim gate As String
  Dim linha As Integer
  qtddecabos = 0
        
  'esse laço percorre toda lista de gates da tabela Geral-Gates
    For i = 1 To (Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("B1").End(xlDown).Row - 1)
        linha = i 'pega a linha do gate
        With Workbooks("taxadeocupacao.xlsm").Sheets("Cabo-Rota").Range("A:XFD")
            gate = Replace(Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Cells(i + 1, 2).Value, "/", "", , 1) 'armazena o valor da célula atual na variável gate
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
                Application.Run "consultardim.calcular_parametros", arraydecabos, linha
                ReDim arraydecabos(0)
                qtddecabos = 0
            End If
        End With
    Next i
    ReDim arraydecabos(0)
End Sub
        
Sub main()
  Application.Calculation = xlCalculationManual
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  excluirPlanilhas
  ReDim arraydecabos(0)
  arraydecabos(0) = ""
  RodarMacroClicked = True
  Application.Run "getalturagate.getaltura"
  Application.Run "getareautilgate.getareautil"
  procurar_cabo_gate
  Application.Run "colorircelulas.colorir"
  Application.Run "hyperlink.testehyperlink"
  
  Application.Calculation = xlCalculationAutomatic
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub
