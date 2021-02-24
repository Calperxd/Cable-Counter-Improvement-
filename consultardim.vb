Dim tempcabos() As String 'armazena os cabos em unidades pro ex: 2x(2x45) + 1x(2x100)
                              'tempcabos(0) = 2x(2x45) primeiro cabo
                              'tempcabos(1) = 1x(2x100) segundo cabo
Dim qtdcabo() As Integer 'armazena as quantidades de cabos pro ex: 2x(2x45) + 1x(2x100)
                              'tempcabos(0) = 2 - primeiro cabo
                              'tempcabos(1) = 1 - segundo cabo
Public areatotal As Double
Dim area As Double


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
Private Sub calcular_parametros(cabos() As String, linhavalue As Integer)
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("I1").Value = "Taxa de ocupação"
    areatotal = 0
    area = 0
    For i = 0 To UBound(cabos)
        Dim dimensaodocabo As String
        dimensaodocabo = Split(cabos(i), "_")(2)
        tempcabos = Split(dimensaodocabo, "+")
        acumuladordeareaatual
        areatotal = areatotal + area
        area = 0
    Next i
    Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("I" & (linhavalue + 1)).Value = (areatotal / Workbooks("taxadeocupacao.xlsm").Sheets("Geral-Gates").Range("H" & (linhavalue + 1)).Value) * 100
    areatotal = 0
    area = 0
    ReDim qtdcabo(0)
End Sub



