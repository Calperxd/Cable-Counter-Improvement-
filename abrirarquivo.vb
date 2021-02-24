Sub UseFileDialogOpen()
    Dim lngCount As Long
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        teste = .Show
        If .Show = -1 Then
            Path = .SelectedItems(1)
        End If
    End With
 
End Sub
'    temp = ActiveWorkbook.Name 'recebe o nome do workbook atual
'    Workbooks.Open Filename:=PathName1(0)
'    TabName1 = ActiveWorkbook.ActiveSheet.Name
'    ControlFile1 = ActiveWorkbook.Name
'    Workbooks(temp).Worksheets("Relatorio-SBFR").Delete
'    Workbooks(ControlFile1).Worksheets(TabName1).Copy After:=Workbooks(temp).Worksheets("Atualizar")
'    Workbooks(ControlFile1).Close SaveChanges:=False
'    Workbooks(temp).Worksheets(TabName1).Activate
'    Sheets(TabName1).Range("A1") = Range("A:A").TextToColumns(DataType:=xlDelimited, ConsecutiveDelimiter:=False, Semicolon:=True)
'    Sheets(TabName1).Range("A1").Value = "NAME OF SITE"
'    Sheets(TabName1).Range("A:XFD").Columns.AutoFit
