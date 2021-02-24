Private Sub testehyperlink()
    With Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates")
         .Hyperlinks.Add Anchor:=.Range("A2:A" & Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("B1").End(xlDown).Row), _
        Address:="", _
        ScreenTip:="", _
        TextToDisplay:=""
    End With
    With Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates")
         .Hyperlinks.Add Anchor:=.Range("B2:B" & Workbooks("taxadeocupacao.xlsm").Worksheets("Geral-Gates").Range("B1").End(xlDown).Row), _
        Address:="", _
        ScreenTip:="", _
        TextToDisplay:=""
    End With
End Sub


    